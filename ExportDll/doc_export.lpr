library doc_export;

{$mode objfpc}{$H+}

uses
  Classes, Graphics, SysUtils, ItemanExport, iteman_fpvectorial,
  docxvectorialwriter, InterfaceBase, Interfaces, fpvectorialpkg;

{thread}var
  ExportIteman: TExportIteman;

type
  TCellSpecInt = record
    Name: PWideChar;
    Value: PWideChar;
  end;

  TTableSpecInt = array of array of TCellSpecInt;
  TTableMatrixInt = array of array of PWideChar;

  function Bytes2Image(var ABytes: TBytes): TBitmap;
  var
    lStream: TMemoryStream;
    l: Int64;
  begin
    Result := nil;
    l := Length(ABytes);

    if l > 0 then
    begin
      Result := TBitmap.Create;

      lStream := TMemoryStream.Create;
      try
        lStream.WriteBuffer(ABytes[0], l);
        lStream.Position := 0;
        Result.LoadFromStream(lStream);
      finally
        lStream.Free;
      end;
    end;
  end;

procedure AddSpecifications(ASpecData: TBytes; const AInputFiles, AOutputFiles : array of PWideChar;
  const ATableSpec: TTableSpecInt;  AIsMultiFiles: Boolean); stdcall;
var
  inputFiles, outputFiles: array of String;
  tableSpec: TTableSpec;
  l1, l2, i, j: Integer;
begin
  if Assigned(ExportIteman) then
  begin
    l1 := Length(AInputFiles) - 1;
    SetLength(inputFiles, l1);
    for i := 0 to l1 - 1 do
      inputFiles[i] := AInputFiles[i];

    l1 := Length(AOutputFiles) - 1;
    SetLength(outputFiles, l1);
    for i := 0 to l1 - 1 do
      outputFiles[i] := AOutputFiles[i];

    SetLength(tableSpec, 0);
    if Assigned(ATableSpec) then
    begin
      l1 := Length(ATableSpec) - 1;
      SetLength(tableSpec, l1);

      for i := 0 to l1 - 1 do
        if Assigned(ATableSpec[i]) then
        try
          l2 := Length(ATableSpec[i]) - 1;
          SetLength(tableSpec[i], l2);
          for j := 0 to l2 - 1 do
          begin
            try
              tableSpec[i, j].Name := ATableSpec[i, j].Name
            except
              tableSpec[i, j].Name := '';
            end;

            try
              tableSpec[i, j].Value := ATableSpec[i, j].Value
            except
              tableSpec[i, j].Value := '';
            end;
          end;
        except
        end;
    end;

    ExportIteman.AddSpecifications(Bytes2Image(ASpecData), inputFiles, outputFiles, tableSpec, AIsMultiFiles);
  end;
end;

procedure AddHeaders(const ATitle: PWideChar; ACombobitData, ALogoData: TBytes; AIsDemo: Boolean); stdcall;
var
  lTitle: String;
begin
  if Assigned(ExportIteman) then
  begin
    lTitle := ATitle;
    ExportIteman.AddHeaders(lTitle, Bytes2Image(ACombobitData), Bytes2Image(ALogoData), AIsDemo);
  end;
end;

procedure AddTable(const ATableName: PWideChar; const ATableMatrix: TTableMatrixInt;
  AFirstRowIsHeader: Boolean = true); stdcall;
var
  tableMatrix: TTableMatrix;
  l1, l2, i, j: Integer;
  lName: String;
begin
  if Assigned(ExportIteman) then
  begin
    lName := ATableName;
    SetLength(tableMatrix, 0);

    if Assigned(ATableMatrix) then
    begin
      l1 := Length(ATableMatrix) - 1;
      SetLength(tableMatrix, l1);

      for i := 0 to l1 - 1 do
        if Assigned(ATableMatrix[i]) then
        try
          l2 := Length(ATableMatrix[i]) - 1;

          SetLength(tableMatrix[i], l2);
          for j := 0 to l2 - 1 do
            if Assigned(ATableMatrix[i, j]) then
              tableMatrix[i, j] := ATableMatrix[i, j]
            else
              tableMatrix[i, j] := '';
        except
        end;
    end;

    ExportIteman.AddTable(lName, tableMatrix, AFirstRowIsHeader);
  end;
end;

procedure AddImage(const AImageTitle: PWideChar; AImageData: TBytes); stdcall;
var
  lTitle: String;
begin
  if Assigned(ExportIteman) then
  begin
    lTitle := AImageTitle;
    ExportIteman.AddImage(lTitle, Bytes2Image(AImageData));
  end;
end;

procedure AddItemByItemDesc(APlotsChecked, ATablesChecked: Boolean); stdcall;
begin
  if Assigned(ExportIteman) then
    ExportIteman.AddItemByItemDesc(APlotsChecked, ATablesChecked);
end;

procedure AddSummary(ASummaryData: TBytes); stdcall;
begin
  if Assigned(ExportIteman) then
    ExportIteman.AddSummary(Bytes2Image(ASummaryData));
end;

procedure AddText(const AText: PWideChar); stdcall;
var
  lText: String;
begin
  if Assigned(ExportIteman) then;
  begin
    lText := AText;
    ExportIteman.AddText(lText);
  end;
end;

function InitExport: Boolean; stdcall;
begin
  try
    ExportIteman := TExportIteman.Create;
    Result := Assigned(ExportIteman);
  except
    Result := False;
    ExportIteman := nil;
  end;
end;

procedure ReInit; stdcall;
begin
  if Assigned(ExportIteman) then
    ExportIteman.ReInit;
end;

procedure SaveToFile(const AFileName: PWideChar); stdcall;
var
  lName: String;
begin
  if Assigned(ExportIteman) then
  begin
    lName := AFileName;
    ExportIteman.SaveToFile(lName);
  end;
end;

procedure DoneExport; stdcall;
begin
  if Assigned(ExportIteman) then
    FreeAndNil(ExportIteman);
end;

exports
  AddHeaders,
  AddImage,
  AddItemByItemDesc,
  AddSpecifications,
  AddSummary,
  AddTable,
  AddText,
  DoneExport,
  InitExport,
  ReInit,
  SaveToFile;

end.

