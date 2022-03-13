unit ItemanExport;

{$mode delphi}{$H+}

interface

uses
  Classes, SysUtils, Graphics,
  //fpvectorialpkg,
  fpvectorial, fpvutils,
  IntfGraphics, FPWriteJPEG,
  docxvectorialwriter;
type
  TCellSpec = record
    Name: string;
    Value: string;
  end;

  TTableSpec = array of array of TCellSpec;
  TTableMatrix = array of array of string;

  { TExportIteman }

  TExportIteman = class
  private
    FPage: TvTextPageSequence;
    FDoc: TvVectorialDocument;
    FCenterCellStyle: TvStyle;
    FHeaderCellStyle: TvStyle;
    FSectionTitleCellStyle: TvStyle;
    FMainParagraphStyle: TvStyle;
    FNormaTextBodyStyle: TvStyle;
    FNormaTextSpanStyle: TvStyle;
    FTableNameStyle: TvStyle;
    procedure InitStyles;
    function NewParagraph(AStyle: TvStyle = nil): TvParagraph;
    function NewCellParagraph(ARow: TvTableRow; AStyle: TvStyle=nil): TvParagraph;



  public
    constructor Create;
    destructor Destroy; override;
    procedure ReInit;
    procedure AddSpecifications(Spec: TBitMap; constref InputFiles, OutputFiles : array of string;
      constref TableSpec: TTableSpec;  IsMultiFiles: Boolean);
    procedure SaveToFile(constref FileName: string);
    procedure AddHeaders(constref Title: string; Combobit, Logo: TBitMap; IsDemo: Boolean);
    procedure AddTable(constref TableName: string; constref TableMatrix: TTableMatrix;
      FirstRowIsHeader: Boolean = true);
    procedure AddImage(constref ImageTitle: string; Image: TBitmap);
    procedure AddItemByItemDesc(PlotsChecked, TablesChecked: Boolean);

    procedure AddSummary(Summary: TBitMap);
    procedure AddText(constref Text: string);
  end;

implementation

uses iteman_fpvectorial;

const
  // Most dimensions in FPVectorial are in mm.  If you want to specify
  // anything in other units, be ready to do the conversion...
  ONE_POINT_IN_MM = 0.35278;

procedure TExportIteman.InitStyles;
begin
  FDoc.AddStandardTextDocumentStyles(vfDOCX);
    // Add our own style for centered paragraphs
  FDoc.AddStandardTextDocumentStyles(vfDOCX);

  FSectionTitleCellStyle := FDoc.AddStyle();
  FSectionTitleCellStyle.MarginLeft := -30;
  FSectionTitleCellStyle.Name := 'SectionTitleCellStyle';
  FSectionTitleCellStyle.Font.Name := 'Arial';
  FSectionTitleCellStyle.SetElements := [sseMarginLeft, spbfFontName];

  FHeaderCellStyle      := FDoc.AddStyle();
  FHeaderCellStyle.Kind := vskTextBody; // This style will be applied to the whole Paragraph
  FHeaderCellStyle.Name := 'HeaderCellStyle';
  FHeaderCellStyle.Font.Name := 'Arial';
  FHeaderCellStyle.Font.Bold:=True;
  FHeaderCellStyle.Font.Size := 10;
  FHeaderCellStyle.Font.Color := TColorToFPColor(clWhite);
  FHeaderCellStyle.Alignment := vsaCenter;
  FHeaderCellStyle.MarginTop := 1 * ONE_POINT_IN_MM;
  FHeaderCellStyle.MarginBottom := 1 * ONE_POINT_IN_MM;
  FHeaderCellStyle.SetElements :=
    [spbfFontSize, spbfFontName, spbfAlignment, sseMarginTop, sseMarginBottom, spbfFontBold, spbfFontColor];

  FCenterCellStyle      := FDoc.AddStyle();
  FCenterCellStyle.Kind := vskTextBody; // This style will be applied to the whole Paragraph
  FCenterCellStyle.Name := 'CenterCellStyle';
  FCenterCellStyle.Font.Name := 'Arial';
  FCenterCellStyle.Font.Size := 10;
  FCenterCellStyle.Alignment := vsaCenter;
  FCenterCellStyle.MarginTop := 1 * ONE_POINT_IN_MM;
  FCenterCellStyle.MarginBottom := 1 * ONE_POINT_IN_MM;
  FCenterCellStyle.SetElements :=
    [spbfFontSize, spbfFontName, spbfAlignment, sseMarginTop, sseMarginBottom];

  FMainParagraphStyle           := FDoc.AddStyle();
  FMainParagraphStyle.Font.Color:= TColorToFPColor(clMaroon);
  FMainParagraphStyle.Font.Name := 'Arial';
  FMainParagraphStyle.Font.Size := 16;
  FMainParagraphStyle.Font.Bold := true;
  FMainParagraphStyle.Font.Italic := true;
  FMainParagraphStyle.Name:= 'MainParagraphStyle';
  FMainParagraphStyle.Kind:= vskTextBody;
  FMainParagraphStyle.SetElements := [spbfFontBold, spbfFontName, spbfFontItalic, spbfFontSize, spbfFontColor];

  FNormaTextBodyStyle             := FDoc.AddStyle();
  FNormaTextBodyStyle.Font.Name   := 'Arial';
  FNormaTextBodyStyle.Font.Size   := 10;
  FNormaTextBodyStyle.Name        := 'NormaTextBodyStyle';
  FNormaTextBodyStyle.Kind        := vskTextBody;
  FNormaTextBodyStyle.SetElements := [spbfFontSize, spbfFontName];

  FNormaTextSpanStyle             := FDoc.AddStyle();
  FNormaTextSpanStyle.Font.Name   := 'Arial';
  FNormaTextSpanStyle.Font.Size   := 10;
  FNormaTextSpanStyle.Name        := 'NormaTextSpanStyle';
  FNormaTextSpanStyle.Kind        := vskTextSpan;
  FNormaTextSpanStyle.SetElements := [spbfFontSize, spbfFontName];

  FTableNameStyle      := FDoc.AddStyle();
  FTableNameStyle.Kind := vskTextBody; // This style will be applied to the whole Paragraph
  FTableNameStyle.Name := 'TableNameStyle';
  FTableNameStyle.Font.Name := 'Arial';
  FTableNameStyle.Font.Size := 11;
  FTableNameStyle.Font.Bold := true;
  FTableNameStyle.Font.Italic := true;
  FTableNameStyle.Alignment := vsaCenter;
  FTableNameStyle.MarginBottom := 0.5;
  FTableNameStyle.SetElements :=
    [spbfFontSize, spbfFontBold, spbfFontName, spbfFontItalic, spbfAlignment, sseMarginBottom];

end;

function TExportIteman.NewParagraph(AStyle: TvStyle): TvParagraph;
begin
  Result := FPage.AddParagraph;
  Result.Style := AStyle;
end;

function TExportIteman.NewCellParagraph(ARow: TvTableRow; AStyle: TvStyle
  ): TvParagraph;
begin
  Result := ARow.AddCell.AddParagraph;
  Result.Style := AStyle;
end;

constructor TExportIteman.Create;
begin
  inherited;
  FDoc := TvVectorialDocument.Create;
  ReInit;
end;

procedure TExportIteman.ReInit;
begin
  FDoc.Clear;
  FPage := FDoc.AddTextPageSequence;
  InitStyles;
end;

destructor TExportIteman.Destroy;
begin
  FDoc.Free;
  inherited Destroy;
end;

procedure TExportIteman.AddSpecifications(Spec: TBitMap; constref InputFiles, OutputFiles : array of string;
  constref TableSpec: TTableSpec;  IsMultiFiles: Boolean);
var
  Table: TvTable;
  Row: TvTableRow;
  i, j: Integer;

begin
  NewParagraph(FMainParagraphStyle).AddText('<#Page#>');
  if Assigned(Spec) then
    NewParagraph(FSectionTitleCellStyle).AddRasterImage.RasterImage := Spec.CreateIntfImage
  else
    NewParagraph(FMainParagraphStyle).AddText('Specifications');
  NewParagraph(FMainParagraphStyle).AddText('');

  if IsMultiFiles then
    NewParagraph(FNormaTextBodyStyle).AddText('A multiple runs file was specified. The Windows paths for the input files are provided below:')
  else
    NewParagraph(FNormaTextBodyStyle).AddText('The Windows paths for the input files used in this analysis were:');
  NewParagraph(FNormaTextBodyStyle).AddText('');
  for i:= 0 to Length(InputFiles)-1 do
   NewParagraph(FNormaTextBodyStyle).AddText(InputFiles[i]);
  NewParagraph(FNormaTextBodyStyle).AddText('');
  if IsMultiFiles then
    NewParagraph(FNormaTextBodyStyle).AddText('A multiple runs file was specified. The Windows paths for the output files are provided below:')
  else
    NewParagraph(FNormaTextBodyStyle).AddText('The Windows paths for the output files produced by this analysis were:');
  NewParagraph(FNormaTextBodyStyle).AddText('');
  for i:= 0 to Length(InputFiles)-1 do
    NewParagraph(FNormaTextBodyStyle).AddText(OutputFiles[i]);
  NewParagraph(FNormaTextBodyStyle).AddText('');

  NewParagraph(FNormaTextBodyStyle).AddText('Table 1 presents the specifications and basic information ' +
                                            'concerning the analysis. This provides important documentation of the setup of the program for historical purposes.');
  NewParagraph(FNormaTextBodyStyle).AddText('');

  NewParagraph(FTableNameStyle).AddText('Table 1: Specifications');

  Table := FPage.AddTable;
  Table.PreferredWidth := Dimension(100, dimPercent);
  Table.Style := FDOc.AddStyle;
  Table.Style.Alignment:= vsaRight;
  Table.Style.SetElements:= [spbfAlignment];

  with Table.Borders do
    begin
      Left.Width:= 0.2;
      Right.Width:= 0.2;
      Top.Width:= 0.2;
      Bottom.Width:= 0.2;
      InsideHoriz.Width:= 0.2;
      InsideVert.Width:= 0.2;
    end;
  Row := Table.AddRow;
  Row.BackgroundColor := RGBToFPColor(24, 55, 81);//(253, 194, 53) is yellow; //192,192,192 is silver
  Row.Header := True;
  NewCellParagraph(Row, FHeaderCellStyle).AddText('Specification');
  NewCellParagraph(Row, FHeaderCellStyle).AddText('Value');
  NewCellParagraph(Row, FHeaderCellStyle).AddText('Specification');
  NewCellParagraph(Row, FHeaderCellStyle).AddText('Value');


  {Write of table spec to outpout}
  for i := 0 to Length(TableSpec)-1 do
  begin
    Row := Table.AddRow;
    for j := 0 to Length(TableSpec[i])-1 do
    begin
      NewCellParagraph(Row, FCenterCellStyle).AddText(TableSpec[i,j].Name);
      NewCellParagraph(Row, FCenterCellStyle).AddText(TableSpec[i,j].Value);
    end;
  end;
  NewParagraph(FNormaTextBodyStyle).AddText('');
end;

procedure TExportIteman.SaveToFile(constref FileName: string);
begin
//  FDoc.WriteToFile(FileName, vfDOCX);
  with TvDOCXVectorialWriterEx.Create do
  try
    WriteToFile(FileName, FDoc);
  finally
    Free;
  end;
end;

procedure TExportIteman.AddHeaders(constref Title: string;
    Combobit, Logo: TBitMap; IsDemo: Boolean);
var
  hs1, hs2, hs3: TvStyle;
begin
  hs1 := FDoc.AddStyle();
  hs1.Alignment := vsaCenter;
  hs1.Font.Color:= TColorToFPColor(clBlack);
  hs1.Font.Name := 'Arial';
  hs1.Font.Size := 28;
  hs1.Name:= 'HeaderStyleTop1';
  hs1.Kind:= vskTextBody;
  hs1.SetElements := hs1.SetElements + [spbfAlignment, spbfFontName, spbfFontSize, spbfFontColor];

  hs2 := FDoc.AddStyle();
  hs2.Alignment := vsaCenter;
  hs2.Font.Color:= TColorToFPColor(clBlack);
  hs2.Font.Name := 'Arial';
  hs2.Font.Size := 20;
  hs2.Name:= 'HeaderStyleTop2';
  hs2.Kind:= vskTextBody;
  hs2.SetElements := hs2.SetElements + [spbfAlignment, spbfFontName, spbfFontSize, spbfFontColor];

  hs3 := FDoc.AddStyle();
  hs3.Alignment := vsaCenter;    //vsaRight
  hs3.Font.Color:= TColorToFPColor(clBlack);
  hs3.Font.Name := 'Arial';
  hs3.Font.Size := 14;
  hs3.Name:= 'HeaderStyleTop3';
  hs3.Kind:= vskTextBody;
  hs3.SetElements := hs3.SetElements + [spbfAlignment, spbfFontBold, spbfFontName,
    spbfFontItalic, spbfFontSize, spbfFontColor];
  if Assigned(Combobit) then
    NewParagraph(FSectionTitleCellStyle).AddRasterImage.RasterImage := ComboBit.CreateIntfImage;
  NewParagraph(hs1).AddText(Title);
  NewParagraph(hs2).AddText('Report created on '+DateTimeToStr(Date));
  NewParagraph(FNormaTextBodyStyle).AddText('');
  NewParagraph(hs3).AddText('Copyright © 2021 - Assessment Systems Corporation');

  if IsDemo then
  begin
    NewParagraph(FNormaTextBodyStyle).AddText('');
    NewParagraph(hs3).AddText('This report was produced by the demo version of Iteman 4.4, which is limited to 100 items and 100 examinees.');
  end;

  NewParagraph(FMainParagraphStyle).AddText('<#Page#>');
  if Assigned(Logo) then
    NewParagraph(FSectionTitleCellStyle).AddRasterImage.RasterImage := Logo.CreateIntfImage
  else
    NewParagraph(FMainParagraphStyle).AddText('Introduction');

  {Write intro to rtf outpout}

  with NewParagraph(FNormaTextBodyStyle) do
    begin
      AddText('This report provides the results of a classical item and test ');
      AddText('analysis by the computer program Iteman Version 4.4 (Assessment Systems Corporation, 2017) for ');
      AddText(Title + '.  ');
      AddText('The output is divided into three sections:  ');
    end;

  with FPage.AddList do
    begin
      Style  := FNormaTextBodyStyle;
      //ListStyle := FDoc.StyleNumberList;
      AddParagraph('   1. Specifications');
      AddParagraph('   2. Summary statistics');
      AddParagraph('   3. Item-by-item results');
    end;
 NewParagraph(FNormaTextBodyStyle).AddText('The statistical output is also recorded in a comma-separated value (CSV) file of the same name. ');

end;

procedure TExportIteman.AddTable(constref TableName: string; constref
  TableMatrix: TTableMatrix; FirstRowIsHeader: Boolean);
var
  Table: TvTable;
  Row: TvTableRow;
  i, j: Integer;
begin
 if TableName <> '' then
   NewParagraph(FTableNameStyle).AddText(TableName);
 if Length(TableMatrix) = 0 then exit;

 Table := FPage.AddTable;
 Table.PreferredWidth := Dimension(100, dimPercent);
 with Table.Borders do
   begin
     Left.Width:= 0.2;
     Right.Width:= 0.2;
     Top.Width:= 0.2;
     Bottom.Width:= 0.2;
     InsideHoriz.Width:= 0.2;
     InsideVert.Width:= 0.2;
   end;
 for i := 0 to Length(TableMatrix)-1 do
   begin
     Row := Table.AddRow;
     if (i = 0) and (FirstRowIsHeader) then
       begin
         Row.BackgroundColor := RGBToFPColor(24, 55, 81);//(253, 194, 53) is yellow; //192,192,192 is silver
         Row.Header := True;
       end;
     for j := 0 to Length(TableMatrix[i])-1 do
       if (i = 0) and (FirstRowIsHeader) then
         NewCellParagraph(Row, FHeaderCellStyle).AddText(TableMatrix[i,j])
       else
         NewCellParagraph(Row, FCenterCellStyle).AddText(TableMatrix[i,j]);
   end;
  NewParagraph(FNormaTextBodyStyle).AddText('');
end;

procedure TExportIteman.AddImage(constref ImageTitle: string; Image: TBitmap);
begin
  NewParagraph(FTableNameStyle).AddText(ImageTitle);
  NewParagraph(FTableNameStyle).AddRasterImage.RasterImage := Image.CreateIntfImage;
end;

procedure TExportIteman.AddItemByItemDesc(PlotsChecked, TablesChecked: Boolean);
begin
 NewParagraph(FMainParagraphStyle).AddText('<#Page#>');
 NewParagraph(FMainParagraphStyle).AddText('Item-by-item results');
 with NewParagraph(FNormaTextBodyStyle) do
   begin
     AddText('The following section presents the item-by-item results of the analysis.  ');
     if PlotsChecked then
        AddText('Each item has several tables and a figure.  The figure, called a quantile plot, shows the'+
           ' proportion of examinees selecting each option, for consecutive segments of'+
           ' the examinees as ranked by score.  The key thing to evaluate in this figure is that the'+
           ' line for the correct answer has a positive slope (goes up from left to right), which'+
           ' means that examinees with higher scores tend to answer correctly more often.'+
           '  Conversely, the lines for the incorrect options, called distractors, should have a'+
           ' negative slope.  Note, however, that the use of a small number of groups (e.g., 3 or fewer) oversimplifies the'+
           ' graph, so that items which are very difficult or very easy (that is, discriminating'+
           ' in only the top or bottom 20% of examinees) might appear to have poor quantile plots'+
           ' and classical statistics.  For such items, item response theory presents significant'+
           ' advantages in analysis.');
   end;
 AddText('');

 if TablesChecked then
    AddText('There are four tables presented for each item.') else
    AddText('There are three tables presented for each item.');
 with FPage.AddList do
   begin
     Style  := FNormaTextBodyStyle;
     //ListStyle := FDoc.StyleNumberList;
     AddParagraph('   1. Item information table: records the information supplied by the control file (or Iteman 3 header) for this item.');
     AddParagraph('   2. Item statistics table: overall item statistics.');
     AddParagraph('   3. Option statistics: detailed statistics for each item, which helps diagnose issues in items with poor statistics.');
     if TablesChecked then
       AddParagraph('   4. Quantile plot data: the values used to create the quantile plot.');
   end;

 AddText('');
 AddText('The item statistics table presents overall item statistics in the first row of numbers. '+
         'The two most important item-level statistics for dichotomously scored (correct/incorrect) items '+
         'are the P value and the point-biserial correlation, which represent the difficulty and discrimination of '+
         'the item, respectively.  For polytomously scored (rating scale or partial credit) items, the difficulty '+
         'is represented by the mean (average) item score, while the discrimination is represented by a Pearson r correlation.');
 AddText('');
 AddText('The P value is the proportion of examinees that answered an item in the keyed direction.  '+
         'P ranges from 0 to 1.  A high value (0.95) means that an item is easy, a low value (0.25) means that the item'+
         ' is difficult. The point-biserial correlation (Rpbis) is a measure of the discriminating, or differentiating'+
         ', power of the item.  Rpbis ranges from -1 to 1.  A negative Rpbis is indicative of a bad item as lower scoring'+
         ' examinees are more likely than higher scoring examinees to respond in the keyed direction.');
 AddText('');
 AddText('For rating scale or partial credit items, the mean item score ranges from the minimum to the maximum'+
         ' of the scale.  For example, if the item has a rating scale of 1 to 5, the possible range for the mean is 1 to 5.'+
         '  The Pearson r is similar to the Rpbis in that it ranges from -1 to 1, with a positive r indicating'+
         ' that the item correlates well with total score.');

 AddText('');
 AddText('The option statistics table presents statistics for each individual option (alternative).'+
         '  The key thing to examine in this portion of the table is that no distractors have'+
         ' a higher Rpbis than the correct answer.  That indicates that higher scoring examinees'+
         ' are selecting the incorrect answer, which therefore might be arguably correct.');

 if TablesChecked then
   begin
    AddText('');
    AddText('The quantile plot data table simply presents the values calculated to create the quantile plot.'+
            '  Because it contains the same information, the quantile plot itself presents a useful picture of the item'+#39+'s'+
            ' performance, but this table can be used to examine that performance in detail to help diagnose possible issues.');
   end;

end;

procedure TExportIteman.AddSummary(Summary: TBitMap);
begin
 NewParagraph(FMainParagraphStyle).AddText('<#Page#>');
 if Assigned(Summary) then
   NewParagraph(FSectionTitleCellStyle).AddRasterImage.RasterImage := Summary.CreateIntfImage
 else
   NewParagraph(FMainParagraphStyle).AddText('Summary statistics');
end;

procedure TExportIteman.AddText(constref Text: string);
begin
  NewParagraph(FNormaTextBodyStyle).AddText(Text);
end;

end.

