unit uProtect;

interface

uses
  System.Types,
  System.Classes,
  System.SysUtils,
  System.StrUtils,
  System.IOUtils,

{$IF DEFINED(Win64) or DEFINED(Win32)}
  System.Win.Registry,
  WinAPI.ShellApi,
  WinAPI.Windows,
{$ENDIF}
  FMX.DialogService; { LCLType; }

const
{$IF DEFINED(Win64) or DEFINED(Win32)}
  cFirstStrIndex = 1;
{$ELSE}
  cFirstStrIndex = 0;
{$ENDIF}

type
  TLicenseStatus = (lcsDemo, lcsActive, lcsWillExpire, lcsExpired);
  TLicenseType = (lctNone, lct30, lct365);

  { TProtect }
  TProtect = class(TObject)
  private
    FKey: string;
    FLic: string;
    FLicenseStatus: TLicenseStatus;
    FLicenseType: TLicenseType;
    FExpirationDays: Integer;
    FAssignCallBack: TNotifyEvent;
    function GetHddSerial: string;
    function GenerateKey: string;
    function CalcCheckSum(AStr: string): string;
    function VerifyCheckSum: Boolean;
  public
    constructor Create;

    class function GetProtect: TProtect;
    procedure DecodeKey(var HddSerial: string; var CreateTime: TDateTime);
    procedure DecodeLicense(var HddSerial: string; var ExpDate: TDate);
    function GetLicenseStatusDes(ASupInfo: Boolean): string;
    function GenerateLicense(AValidUntil: TDate): string;
    procedure AssignLicense;

    property Key: string read FKey {$IFDEF GENERATOR} write FKey {$IFEND};
    property License: string read FLic write FLic;
    property LicenseStatus: TLicenseStatus read FLicenseStatus {$IFDEF GENERATOR} write FLicenseStatus {$IFEND};
    property AssignCallBack: TNotifyEvent read FAssignCallBack write FAssignCallBack;
    property LicenseType: TLicenseType read FLicenseType write FLicenseType;
  end;

implementation

const
  REG_KEY = '\software\ass\iteman';
  SEP = '-';
  RAND_VAL: array [1 .. 10] of string = ('0', '1', '2', '3', '4', '5', '6', '7', '8', '9');

var
  oProtect: TProtect;

  { TProtect }

constructor TProtect.Create;
var
  reg: TRegistry;
  // s1: string;
  // t: TDateTime;
begin
  FLicenseStatus := lcsDemo;
  FLicenseType := lctNone;
  FExpirationDays := -1;
  FKey := '';
  FLic := '';

{$IFNDEF GENERATOR}
{$IF DEFINED(Win64) or DEFINED(Win32)}
  reg := TRegistry.Create();
  try
    // the registry root is HKCU to avoid rights access problems
    reg.RootKey := NativeUInt($80000001); // HKEY_CURRENT_USER;
    if (reg.OpenKey(REG_KEY, True)) then
    begin
      FKey := reg.ReadString('key'); // reading or generating the key

      if (FKey = '') then
      begin
        FKey := Self.GenerateKey;
        reg.WriteString('key', FKey);
      end;

      FLic := reg.ReadString('lic'); // reading the licence
    end
    else
      raise Exception.Create('Error reading Windows Registry.' + #13#10 + 'Is not possible to load license data.');
  finally
    FreeAndNil(reg);
    AssignLicense;
  end;
{$ENDIF}
{$ENDIF}
end;

class function TProtect.GetProtect: TProtect;
begin
  Result := oProtect;
end;

function TProtect.GetLicenseStatusDes(ASupInfo: Boolean): string;
begin
  case FLicenseStatus of
    lcsDemo:
      Result := 'Demo License';
    lcsActive:
      Result := 'Active License';
    lcsWillExpire:
      Result := Format('The licence will expire in %d days', [FExpirationDays]);
    lcsExpired:
      Result := 'License Expired';
  end;

  if (ASupInfo) then
    case FLicenseStatus of
      lcsDemo:
        Result := Result + ' (click here to register)';
      lcsWillExpire, lcsExpired:
        Result := Result + ' (click here to renew)';
    end;
end;

function TProtect.GetHddSerial: string;
var
  lSysDriveName: string;
  lVolumeSerialNumber: DWORD;
  lMaximumComponentLength: DWORD;
  lFileSystemFlags: DWORD;
  lSerialNumber: string;
begin
{$IF DEFINED(Win64) or DEFINED(Win32)}
  lSysDriveName := System.IOUtils.TPath.GetHomePath();
  lSysDriveName := System.IOUtils.TPath.GetPathRoot(lSysDriveName);

  // get the serial number of the volume SysDrive. It may not be the leter "c:\"
  GetVolumeInformation(pWideChar(ExtractFileDir(lSysDriveName)), nil, 0, @lVolumeSerialNumber, lMaximumComponentLength,
    lFileSystemFlags, nil, 0);
  lSerialNumber := IntToHex(HiWord(Integer(lVolumeSerialNumber)), 4) +
    IntToHex(LoWord(Integer(lVolumeSerialNumber)), 4);
{$ENDIF}
  Result := lSerialNumber;
end;

function TProtect.GenerateKey: string;
var
  sNowP, sHddSerialP, sNow, sHddSerial: string;
  iCount: Integer;
begin
  // add a separator to avoid errors with differnt string serial length
  sHddSerial := GetHddSerial + SEP;

  // the time is converted in the numeric value with out the decimal sepatator
  // the result string is reversed
  sNow := ReverseString(StringReplace(FloatToStr(Now), FormatSettings.DecimalSeparator, '', [rfReplaceAll]));
  Result := '';
  iCount := 0;

  // the key will be generated with alternate sub string from both the values
  repeat
    Inc(iCount);
    sNowP := Copy(sNow, (iCount * 2) - 1, 2);
    sHddSerialP := Copy(sHddSerial, (iCount * 2) - 1, 2);
    Result := Result + sNowP + sHddSerialP;
  until (sNowP = '') and (sHddSerialP = '');
end;

procedure TProtect.AssignLicense;
var
  sSerial: string;
  dtNow, dtExp: TDate;
  reg: TRegistry;
begin
  FLicenseStatus := lcsDemo;
  if (FLic > '') then
  begin
    Self.DecodeLicense(sSerial, dtExp);
    dtNow := Trunc(Now);

    if (sSerial <> GetHddSerial) or (dtExp < dtNow) then
      FLicenseStatus := lcsExpired
    else if (dtExp < (dtNow + 14)) then
      FLicenseStatus := lcsWillExpire
    else
      FLicenseStatus := lcsActive;

    if FLicenseStatus in [lcsWillExpire, lcsActive] then
    begin
{$IF DEFINED(Win64) or DEFINED(Win32)}
      reg := TRegistry.Create();
      reg.RootKey := NativeUInt($80000001); // HKEY_CURRENT_USER;
      try
        if (reg.OpenKey(REG_KEY, True)) then
          reg.WriteString('lic', FLic);
      finally
        FreeAndNil(reg);
      end;
{$ENDIF}
    end;
  end;

  if Assigned(FAssignCallBack) then
    FAssignCallBack(Self);
end;

function TProtect.CalcCheckSum(AStr: string): string;
var
  iTot: Integer;
  iCount: Integer;
begin
  // calc a sum of the license
  iTot := 0;
  for iCount := 1 to Length(AStr) do
    iTot := iTot + (Ord(AStr[iCount]) * 19);

  // calc the checksum
  Result := IntToStr((98 - ((iTot * 100) mod 97)) mod 97);

  // the checksum MUST be 2 cheracters
  if (Length(Result) = 1) then
    Result := '0' + Result;
end;

function TProtect.VerifyCheckSum: Boolean;
var
  iTot: Integer;
  iCount: Integer;
  sStr: string;
  sCheck: string;
begin
  Result := False;
  try
    sStr := Copy(FLic, 1, Length(FLic) - 2);
    sCheck := Copy(FLic, Length(FLic) - 1, 2);
    // calc a sum of the license
    iTot := 0;
    for iCount := 1 to Length(sStr) do
    begin
      iTot := iTot + (Ord(sStr[iCount]) * 19);
    end;
    // calc the validation code
    iTot := StrToInt(IntToStr(iTot) + sCheck) mod 97;
    Result := (iTot = 1);
  except
  end;
end;

procedure TProtect.DecodeKey(var HddSerial: string; var CreateTime: TDateTime);
var
  sTmp, sPart, sTime, sSerial: string;
  iIndex, iCount: Integer;
begin
  HddSerial := '';
  CreateTime := 0;
  iCount := 0;
  sTime := '';
  sSerial := '';
  // where is the hdd serial separator?
  iIndex := Pos(SEP, FKey);

  // all the sub string after the separator is part of the date time string
  // i'm going to extract only the part before the separator
  sTmp := Copy(FKey, 1, iIndex - 1);
  repeat
    Inc(iCount);
    // the key was created with alternate parts of the strings
    // i'm going to extract the part and reassemble the real values
    sPart := Copy(sTmp, (iCount * 2) - 1, 2);
    if Odd(iCount) then
      sTime := sTime + sPart
    else
      sSerial := sSerial + sPart;
  until (sPart = '');

  // convert the numeric date time to a TDateTime (that is a real number...)
  sTime := ReverseString(sTime + Copy(FKey, iIndex + 1, Length(FKey)));
  sTime := Copy(sTime, 1, 5) + FormatSettings.DecimalSeparator + Copy(sTime, 6, Length(sTime));

  // serial
  HddSerial := sSerial;
  CreateTime := StrToFloatDef(sTime, 0);
end;

procedure TProtect.DecodeLicense(var HddSerial: string; var ExpDate: TDate);
var
  sNoCheck, sTmp, sPart, sDate, sSerial: string;
  iPos, iIndex, iCount: Integer;
begin
  HddSerial := '';
  ExpDate := 0;
  iCount := 0;
  sDate := '';
  sSerial := '';

  // first of all i'm going to veify the checksum
  if (Self.VerifyCheckSum) then
  begin
    // the last two digits are the checksum
    sNoCheck := Copy(FLic, 1, Length(FLic) - 2);

    // where is the hdd serial separator?
    iIndex := Pos(SEP, sNoCheck);

    // all the sub string after the separator is part of the date time string
    // i'm going to extract only the part before the separator
    sTmp := Copy(sNoCheck, 1, iIndex - 1);
    iPos := 1;

    repeat
      Inc(iCount);
      // the key was created with alternate parts of the strings
      // i'm going to extract the part and reassemble the real values
      if Odd(iCount) then
      begin
        sPart := Copy(sTmp, iPos, 2);
        sDate := sDate + sPart;
        iPos := iPos + 2;
      end
      else
      begin
        sPart := Copy(sTmp, iPos, 3);
        sSerial := sSerial + sPart;
        iPos := iPos + 3;
      end;
    until (sPart = '');

    // convert the numeric date time to a TDate (that is a number...)
    sDate := sDate + Copy(sNoCheck, iIndex + 1, Length(sNoCheck));
    // the second and 5th characters was recalculated as the ord of the char
    // the result of the ord is a 2 digits number, the values are:
    // 2nd = 2..3     5th = 6..7
    sDate := sDate[1] + Char(StrToInt(sDate[2] + sDate[3])) + sDate[4] + sDate[5] + Char(StrToInt(sDate[6] + sDate[7]))
      + Copy(sDate, 8, Length(sDate));
    // i added a random number at pos 1 and 4, now i'm going to delete that characters
    sDate := sDate[2] + sDate[3] + Copy(sDate, 5, Length(sDate));
    // reverse again
    sDate := ReverseString(sDate);
  end
  else
  begin
{$IFDEF GENERATOR}
    TDialogService.ShowMessage('The license key is invalid!');
{$IFEND};
  end;

  // set the result
  HddSerial := sSerial;
  ExpDate := StrToFloatDef(sDate, 0);
end;

function TProtect.GenerateLicense(AValidUntil: TDate): string;
var
  sExpP: string;
  sExp: string;
  sSerialP: string;
  sSerial: string;
  dtKeyCreated: TDateTime;
  iCount: Integer;
begin
  Randomize;
  Result := '';

  Self.DecodeKey(sSerial, dtKeyCreated);

  if (sSerial > '') and (dtKeyCreated > 0) then
  begin
    sExp := ReverseString(IntToStr(Trunc(Int(AValidUntil))));
    // date in string without the decimal part, reversed
    sExp := RandomFrom(RAND_VAL) + Copy(sExp, 1, 2) + RandomFrom(RAND_VAL) + Copy(sExp, 3, Length(sExp));
    // i add a random number at pos 1 and 3
    sExp := sExp[1] + IntToStr(Ord(sExp[2])) + sExp[3] + sExp[4] + IntToStr(Ord(sExp[5])) + Copy(sExp, 6, Length(sExp));
    // the second and 5th characters are recalculated as the ord of the char
    sSerial := sSerial + SEP; // add the serial separator
    iCount := 0; // mix the values

    repeat
      Inc(iCount);
      sExpP := Copy(sExp, (iCount * 2) - 1, 2);
      sSerialP := Copy(sSerial, (iCount * 3) - 2, 3);
      Result := Result + sExpP + sSerialP;
    until (sExpP = '') and (sSerialP = '');

    Result := Result + Self.CalcCheckSum(Result);
  end
  else
    TDialogService.ShowMessage('The key is invalid.');
end;

initialization

oProtect := TProtect.Create;

finalization

FreeAndNil(oProtect)

end.
