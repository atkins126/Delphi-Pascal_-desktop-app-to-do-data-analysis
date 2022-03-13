unit uItemanExport;

interface

uses
  Fmx.Graphics;

type
  TCellSpec = record
    Name: String;
    Value: String;
  end;

  TTableSpec = array of array of TCellSpec;
  TTableMatrix = array of array of String;

const
  ExportLibName = 'doc_export.dll';

procedure AddSpecifications(ASpecData: TBitmap; const AInputFiles, AOutputFiles : array of PWideChar;
  const ATableSpec: TTableSpec;  AIsMultiFiles: Boolean);
procedure AddHeaders(const ATitle: String; ACombobitData, ALogoData: TBitmap; AIsDemo: Boolean);
procedure AddTable(const ATableName: String; const ATableMatrix: TTableMatrix;
  AFirstRowIsHeader: Boolean = true);
procedure AddImage(const AImageTitle: String; AImage: TBitmap);
procedure AddSummary(ASummaryData: TBitmap);
procedure AddItemByItemDesc(APlotsChecked, ATablesChecked: Boolean); stdcall; external ExportLibName;
procedure AddText(const AText: PWideChar); stdcall; external ExportLibName;
function InitExport: Boolean; stdcall; external ExportLibName;
procedure DoneExport; stdcall; external ExportLibName;
procedure ReInit; stdcall; external ExportLibName;
procedure SaveToFile(const AFileName: PWideChar); stdcall; external ExportLibName;

implementation

uses SysUtils, Classes, FMX.Consts, FMX.Surfaces;

type
  TCellSpecInt = record
    Name: PWideChar;
    Value: PWideChar;
  end;

  TTableSpecInt = array of array of TCellSpecInt;
  TTableMatrixInt = array of array of PWideChar;

procedure AddTableInt(const ATableName: PWideChar; const ATableMatrix: TTableMatrixInt;
  AFirstRowIsHeader: Boolean = true); stdcall; external ExportLibName name 'AddTable';
procedure AddSpecificationsInt(ASpecData: TBytes; const AInputFiles, AOutputFiles : array of PWideChar;
  const ATableSpec: TTableSpecInt;  AIsMultiFiles: Boolean); stdcall; external ExportLibName name 'AddSpecifications';
procedure AddImageInt(const AImageTitle: PWideChar; AImageData: TBytes); stdcall; external ExportLibName name 'AddImage';
procedure AddHeadersInt(const ATitle: PWideChar; ACombobitData, ALogoData: TBytes;
  AIsDemo: Boolean); stdcall; external ExportLibName name 'AddHeaders';
procedure AddSummaryInt(ASummaryData: TBytes); stdcall; external ExportLibName name 'AddSummary';


function Image2Bytes(var AImage: TBitmap): TBytes;
var
  FStream: TMemoryStream;
  Surf: TBitmapSurface;
begin
  Result := nil;

  if Assigned(AImage) then
  begin
    FStream := TMemoryStream.Create;

    try
      TMonitor.Enter(AImage);
      try
        Surf := TBitmapSurface.Create;
        try
          Surf.Assign(AImage);
          TBitmapCodecManager.SaveToStream(FStream, Surf, SBMPImageExtension);
        finally
          Surf.Free;
        end;
      finally
        TMonitor.Exit(AImage);
      end;

      FStream.Position := 0;
      SetLength(Result, FStream.Size);
      FStream.Read(Result, FStream.Size);
    finally
      FreeAndNil(FStream);
    end
  end;
end;

procedure AddSummary(ASummaryData: TBitmap);
begin
  AddSummaryInt(Image2Bytes(ASummaryData));
end;

procedure AddImage(const AImageTitle: String; AImage: TBitmap);
var
  FTitle: PWideChar;
begin
  FTitle := PWideChar(AImageTitle);
  AddImageInt(FTitle, Image2Bytes(AImage));
end;

procedure AddHeaders(const ATitle: String; ACombobitData, ALogoData: TBitmap; AIsDemo: Boolean);
var
  FTitle: PWideChar;
begin
  FTitle := PWideChar(ATitle);
  AddHeadersInt(FTitle, Image2Bytes(ACombobitData), Image2Bytes(ALogoData), AIsDemo);
end;

procedure AddSpecifications(ASpecData: TBitmap; const AInputFiles, AOutputFiles : array of PWideChar;
  const ATableSpec: TTableSpec;  AIsMultiFiles: Boolean);
var
  TableSpecInt: TTableSpecInt;
  i, j, len1, len2: Integer;
begin
  len1 := Length(ATableSpec);
  SetLength(TableSpecInt, len1);

  for i := 0 to len1 - 1 do
  begin
    len2 := Length(ATableSpec[i]);
    SetLength(TableSpecInt[i], len2);

    for j := 0 to len2 - 1 do
    begin
      if not ATableSpec[i, j].Name.IsEmpty then
        TableSpecInt[i, j].Name := PWideChar(ATableSpec[i, j].Name)
      else
        TableSpecInt[i, j].Name := '';

      if not ATableSpec[i, j].Value.IsEmpty then
        TableSpecInt[i, j].Value := PWideChar(ATableSpec[i, j].Value)
      else
        TableSpecInt[i, j].Value := '';
    end;
  end;

  AddSpecificationsInt(Image2Bytes(ASpecData), AInputFiles, AOutputFiles, TableSpecInt, AIsMultiFiles);
end;

procedure AddTable(const ATableName: String; const ATableMatrix: TTableMatrix;
  AFirstRowIsHeader: Boolean = true);
var
  TableMatrixInt: TTableMatrixInt;
  tableName: String;
  pName: PWideChar;
  i, j, len1, len2: Integer;
begin
  if ATableName.IsEmpty then
    tableName := ''
  else
    tableName := ATableName;

  len1 := Length(ATableMatrix);
  SetLength(TableMatrixInt, len1);
  for i := 0 to len1 - 1 do
  begin
    len2 := Length(ATableMatrix[i]);
    SetLength(TableMatrixInt[i], len2);

    for j := 0 to len2 - 1 do
      if not ATableMatrix[i, j].IsEmpty then
        TableMatrixInt[i, j] := PWideChar(ATableMatrix[i, j])
      else
        TableMatrixInt[i, j] := '';
  end;

  pName := PWideChar(tableName);
  AddTableInt(pName, TableMatrixInt, AFirstRowIsHeader);
end;

end.
