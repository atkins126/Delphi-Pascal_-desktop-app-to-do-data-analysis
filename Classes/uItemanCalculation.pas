unit uItemanCalculation;

{$RANGECHECK OFF} // TODO : enable range checking in this unit back on


interface

uses
  System.SysUtils,
  System.Types,
  System.UITypes,
  System.Classes,
  System.Variants,
  System.Rtti,
  System.DateUtils,
  System.Math,
  uItemanExport,
  FMX.Forms,
  FMX.Objects,
  FMX.Graphics,
  FMX.Types,
  StrUtils;

type
  TIntegerArray = Array of Array of Integer;
  TRealArray = Array of Array of Real;
  TStringArray = Array of String;
  TChar1Array = Array of Char;
  TChar2Array = Array of Array of Char;

  TBitmapHelper = class helper for FMX.Graphics.TBitmap
  public
    procedure LoadFromResource(AResourceName: String; AWidth: Integer = 0; AHeight: Integer = 0);
  end;

  TItemanCalculation = class
  private
    FGraphBitmap: TBitmap;
    FTitleBitmap: TBitmap;
    FIntroductionBitmap: TBitmap;
    FSpecBitmap: TBitmap;
    FSummBitmap: TBitmap;

    FUseMRF: Boolean;
    FOutputDir: string;

    FNumMult: array [1 .. 4] of Integer; // scored,all,pretest,not included
    FNumRate: array [1 .. 4] of Integer;
    FDomMult: array [1 .. 100] of Integer;
    FDomRate: array [1 .. 100] of Integer;

    FMaxP, FMinP, FMaxR, FMinR, FMaxMean, FMinMean: Real;
    FNumFlagItems: Integer;
    FOmitChar: Char;
    FNAChar: Char;
    FNumberOfExaminees: Integer;
    FNumCut: Integer;
    FAllOk: Boolean;
    FRespOK: Boolean;

    FBBOindexArray: array of array of Real;
    FBBOFlags: array of Integer;

    FNumberItemsSoFar: Integer;
    FScoredItems: Integer;
    FPreTest: Integer;

    FNumberOfItems: Integer;
    FValidItems: Integer;
    FNumDomains: Integer;

    FStoreDomain: array [1 .. 50] of string;
    IsZero: Boolean;
    FNumberForDif: Integer;

    FErrorMessage: string;

    Fdata: string;
    Fctrl: string;
    Foutp: string;
    Fmrf1: string;
    Fmaxcategory: Integer;
    // FExport: TExportIteman;
    // FItem: Integer;
    FMaxItemOptions: Integer;
    FExternalScoreArray: array [1 .. 500000] of Real;
    FDomainStatsArray: array [1 .. 100, 1 .. 16] of Real;
    FTestStatsArray: array [1 .. 5, 1 .. 16] of Real;
    FGroupCounts: array [1 .. 7] of Integer;
    FScaledScoreArray: array [1 .. 500000] of Real;
    FScoreSS: Real;
    FOmch: Char;
    FOptionCountsArray: array [1 .. 10000, 1 .. 17] of Integer;
    FOptionPropsArray: array [1 .. 10000, 1 .. 17] of Real;
    FOptionRpbisArray: array [1 .. 10000, 1 .. 17] of Real;
    FOptionRbisArray: array [1 .. 10000, 1 .. 17] of Real;
    FOptionMeansArray: array [1 .. 10000, 1 .. 17] of Real;
    FOptionSDArray: array [1 .. 10000, 1 .. 17] of Real;
    FMinItem1: Real;
    FMaxItem1: Real;
    FMinRPBis: Real;
    FMinCorr: Real;
    FItemStatsArray: array [1 .. 10000, 1 .. 8] of Real;
    FCh: Char;
    FKey: Integer;
    FItemN: Integer;
    FResponseAsInt: Integer;
    FRespAsInt: Boolean;
    FWts: Integer;
    FRpbisCalcArray: array [1 .. 99, 1 .. 3] of Real;
    FScoreSum: Real;
    FSpurN: Real;
    FSpurScoreMean: Real;
    FP: Real;
    FZ: Real;
    FSpurScoreSD: Real;
    FRpbisSum: Real;
    FOrdinate: Real;
    FNforRpbis: Integer;
    FIsMult: Boolean;
    FIsRate: Boolean;
    FScoreMean: Real;
    FVarianceForKR20Without: Real;
    FDomainArray: array [1 .. 500000, 1 .. 50] of Integer;
    FDomainCorrArray: array [1 .. 50, 1 .. 50] of Real;
    FDomainCS: Real;
    FDomScaleStat: array [1 .. 100, 1 .. 4] of Real;
    FMaxCsem: Real;
    FMinCsem: Real;
    FIX: LongInt;
    FIY: LongInt;
    FIZ: LongInt;
    FNumUniqueFreq: Integer;
    FEICArray: array of array of Integer; // used to be local now global
    FEEICArray: array of array of Integer; // used to be local now global
    FProppassing: Real;
    FFlagID: array [1 .. 10000] of Integer;
    FFlagArray: array of string;
    // columns 6&7 now LM HM LowMean HighMean, column 8 is DIF
    FSubGroupPArray: array [1 .. 105] of Real;
    FGroup: Integer;
    FPSum: Real;
    FPTemp: Real;
    nbin, tempint, Start, Finish, Domain, brange, summarytablerows, ExamineeIDColumns: Integer;
    scoredrate: Boolean;
    GroupedFreqArray: Array [1 .. 21, 1 .. 4] of Real;
    rawcut: Real;

    function CharCode(aChar: Char): Integer;
    function GetStrFromCharArr(aCharsArray: TChar1Array): string;
    function IsNumeric(FCh: Char): Boolean;
    function IsAlphaBetic(FCh: Char): Boolean;
    function ConvertResponseToInteger(FCh: Char; iteminfoarray: TChar2Array; FItem, nalt: Integer): Integer;
    function RespToInt(FCh: Char; iteminfoarray: TChar2Array; FItem: Integer; itemscalearray: TIntegerArray): Integer;
    procedure mantelH(item: Integer; responsearray, iteminfoarray: TChar2Array; difrankarray: TIntegerArray;
      difarray: TRealArray);
    function cdfn(x: single): single;
    function erfn(x: single): single;
    procedure DIFgroups(difrankarray, scorearray: TIntegerArray; iteminfoarray, responsearray: TChar2Array);
    procedure ScoreGroups(RankingArray, ReduceArray, scorearray: TIntegerArray;
      iteminfoarray, responsearray: TChar2Array);
    procedure SumStats(iteminfoarray: TChar2Array; scorearray: TIntegerArray);
    procedure OptionStats(iteminfoarray, responsearray: TChar2Array;
      itemscalearray, nkeyarray, scorearray: TIntegerArray);
    function cdfni(p: single): single;
    procedure ScaleStats(iteminfoarray: TChar2Array);
    procedure EstimateKR20Without(responsearray, iteminfoarray: TChar2Array;
      itemscalearray, nkeyarray, scorearray: TIntegerArray);
    procedure DomainStats(iteminfoarray, responsearray: TChar2Array; itemscalearray: TIntegerArray;
      domscalearray: TRealArray; nkeyarray: TIntegerArray);
    function EstimateKR20(Domain: Integer; iteminfoarray: TChar2Array; itemscalearray: TIntegerArray): Real;
    function CalculateMeanP(Domain: Integer; iteminfoarray: TChar2Array; itemscalearray: TIntegerArray): Real;
    function CalculateItemMean(Domain: Integer; iteminfoarray: TChar2Array; itemscalearray: TIntegerArray): Real;
    function CalculateMeanRpbis(Domain: Integer; iteminfoarray: TChar2Array; itemscalearray: TIntegerArray): Real;
    procedure Reliability(splitscore1, splitscore2, splitdomain1, splitdomain2, itemscalearray: TIntegerArray;
      iteminfoarray, responsearray: TChar2Array);
    procedure CSem(itemscalearray: TIntegerArray; iteminfoarray, responsearray: TChar2Array; csemarray: TRealArray);
    function RandomNum(counter: Integer): Real;
    procedure ShellSortInt(numericarray: TIntegerArray; n: Integer);
    procedure RawFreq(rawfrequency, scorearray: TIntegerArray);
    procedure BBOindex(iteminfoarray, responsearray: TChar2Array; itemscalearray: TIntegerArray);
    function Factorial(Num: LongInt): LongInt;
    function KeyAsInt(iteminfoarray: TChar2Array; item: Integer): Integer;
    procedure SubGroups(aItemInfoArray, aResponseArray: TChar2Array; aOmch: Char; aI: Integer;
      aRankIngArray, aReduceArray, aItemsCaleArray: TIntegerArray);

    procedure CalcGF(Min, Max: Real; call: Integer; scorearray: TIntegerArray);
    procedure CalcItemGF(Min, Max: Real; call, denom: Integer; type1: Char; iteminfoarray: TChar2Array);
    function GetTableSpec: TTableSpec;
    procedure PrintFreq10(const prec: string; Min, Max: Real);
    procedure PrintFreq15(const prec: string; Min, Max: Real);
    procedure PrintFreq20(const prec: string; Min, Max: Real);
    procedure PrintFirstSummary(prec: string; iteminfoarray: TChar2Array; itemidarray: TStringArray;
      itemscalearray: TIntegerArray; difarray: TRealArray);
    procedure PrintItemTables(const prec: string; item: Integer; iteminfoarray: TChar2Array; itemidarray: TStringArray;
      itemscalearray, nkeyarray: TIntegerArray; start0: Char; difarray: TRealArray);
    function Livingston: Real;
    procedure DrawDistGraph(Min, Max: Real; nit, nunit, call: Integer; graph: TBitmap; title: string);
    procedure DrawScatterplot(Min, Max: Real; nit, nunit, call: Integer; graph: TBitmap; iteminfoarray: TChar2Array;
      type1: Char; title: string);
    procedure DrawCsemGraph(nit: Integer; graph: TBitmap; title: string; csemarray: TRealArray);
    procedure DrawItemGraph(graph: TBitmap; item: Integer; itemscalearray: TIntegerArray; itemidarray: TStringArray;
      key1, start0: Char);
    function GetFloatValue(aFloatValueInStr: string = '0'; aFloatDidgets: Integer = -3): Extended;
    function GetColumnsCountOfCSVFile(Filename: string): Integer;
  public
    constructor Create(const AFlags: array of String); reintroduce;
    destructor Destroy; override;

    function RunIteman(const AItemResponsesBeginInColumn, ADataMatrixVariableAdditionalColumns,
      AItemControlVariableAdditionalRows: Integer): Boolean;

    property UseMRF: Boolean read FUseMRF write FUseMRF;
    property OutputDir: string read FOutputDir write FOutputDir;

    // Output error message
    property ErrorMessage: string read FErrorMessage write FErrorMessage;

    { Calculation properties }
    property data: string read Fdata write Fdata;
    property ctrl: string read Fctrl write Fctrl;
    property outp: string read Foutp write Foutp;
    property mrf1: string read Fmrf1 write Fmrf1;
    property maxcategory: Integer read Fmaxcategory write Fmaxcategory;
  end;

implementation

uses
  ufrmMain;

const
  C_Pict_Width = 794;

  { TItemanCalculation }

constructor TItemanCalculation.Create(const AFlags: array of String);
var
  i: Integer;
begin
  inherited Create;

  SetLength(FFlagArray, Length(AFlags) + 1);
  for i := Low(AFlags) to High(AFlags) do
    FFlagArray[i + 1] := AFlags[i];

  FGraphBitmap := TBitmap.Create;
  FTitleBitmap := TBitmap.Create;
  FTitleBitmap.LoadFromResource('Title_page_image', C_Pict_Width);
  FIntroductionBitmap := TBitmap.Create;
  FIntroductionBitmap.LoadFromResource('Introduction_header', C_Pict_Width);
  FSpecBitmap := TBitmap.Create;
  FSpecBitmap.LoadFromResource('Specifications_header', C_Pict_Width);
  FSummBitmap := TBitmap.Create;
  FSummBitmap.LoadFromResource('Summary_statistics_header', C_Pict_Width);
end;

destructor TItemanCalculation.Destroy;
begin
  FGraphBitmap.Free;
  FTitleBitmap.Free;
  FIntroductionBitmap.Free;
  FSpecBitmap.Free;
  FSummBitmap.Free;
  SetLength(FFlagArray, 0);
  inherited;
end;

function TItemanCalculation.RunIteman(const AItemResponsesBeginInColumn, ADataMatrixVariableAdditionalColumns,
  AItemControlVariableAdditionalRows: Integer): Boolean;
const
  DataMatrixAdditionalColumns = 2; // two columns which always are there - Group	and Inclusion Code
var
  nreport, itemcap, ticker, start1, idread, tabcounter, commacounter, mrint, pcint: Integer;
  k, { avc, } lastvalid, vv, previd, i, j, pr, row, nex, errno, numkey, percentile, figureno, tableno, SCO: Integer;
  temp1, isdocx, fino, PreStr, PreText, prec, path, Filename, mrfname: string;
  matchwarning, revmatchwarning, contint: string;
  Rrange, x: Real;
  datacontrolerror: array [1 .. 1200] of Boolean;
  patharray: array [1 .. 500] of Char;
  readformat: array [1 .. 50] of Char;
  delim1, controldelim: Char;
  lCh: Char;
  DomainMean, DomainSD, DomainSS, DomainCS, DomainSum, minrpbis, mincorr: Real;
  Buffer, fl1, fl2, fl3, fl4, fl5, newdocxoutput: string;

  lResponseArray: TChar2Array;
  lItemIDArray: TStringArray;
  lExamineeIDArray: TStringArray;
  lItemInfoArray: TChar2Array;
  lItemScaleArray: TIntegerArray;
  lRankingArray: TIntegerArray;
  lReduceArray: TIntegerArray;
  lSplitScore1: TIntegerArray;
  lSplitScore2: TIntegerArray;
  lDomScaleArray: TRealArray;
  lCsemArray: TRealArray;
  lSplitDomain1: TIntegerArray;
  lSplitDomain2: TIntegerArray;
  lNKeyArray: TIntegerArray;
  lDifRankArray: TIntegerArray;
  lScoreArray: TIntegerArray;
  lDifarray: TRealArray;
  lRawFrequency: TIntegerArray;

  tbMatrix: TTableMatrix;
  tmp: string;

  lIsDelim: Boolean;
  lMismatch: Boolean;
  lItemIsFlagged: Boolean;
  lCheckKey: Boolean;
  lInputOk: Boolean;
  lColumnOk: Boolean;
  lTypeOk: Boolean;
  lAltOk: Boolean;
  lDomOk: Boolean;
  lEndLine: Boolean;
  lZeroOk: Boolean;
  lKeyOk: Boolean;

  lAllPilot: Boolean;
  lAllPC: Boolean;
  scoredrate: Boolean;
  lScoredMult: Boolean;

  lMRFFile: TextFile;
  lDataMatrixFile: TextFile;
  lItemInfoFile: TextFile;
  lOutputFile: TextFile;
  lScoresFile: TextFile;
  lScoredMatrixFile: TextFile;
  lExternalScoreFile: TextFile;

  lChar: Char;

  lExamineeIDColumns: Integer;
  lNumberOfEntries: Integer;

  lCSVFileName: string;
  lRunTitle: string;
  lOmitString: string;
  lNAString: string;

  lFileRow: Integer;
  lIDBegin: Integer;
  lColBox: Integer;
  lGroupCodeColumn: Integer;
  lIDSpace: Integer;
  lTabChar: Char;
  lIDsArray: TChar1Array; // string[100];
  lTempInt: Integer;
  lSameDomain: Boolean;
  lReadIDArray: array [1 .. 200] of Char; { Only works when static }
  lOutputFileName: string;
  lControlText: TextFile;
  lMaxExamId: Integer;
  lExamineeID: string;
  lExternalScore: Real;
  lItem: Integer;

  lRealTest: Real;

  function CheckSettings(): Boolean;
  begin
    Result := True;

    if not UseMRF then
    begin
      if frmMain.eDataMatrixFile.Text = '' then
      begin
        ErrorMessage := ('Please select an data matrix file.');
        lInputOk := False;
        Result := False;
        Exit;
      end;

      if (frmMain.eItemControlFile.Text = '') and (frmMain.cbDataMatrixFile.IsChecked = False) then
      begin
        ErrorMessage := ('Please select an item control file.');
        lInputOk := False;
        Result := False;
        Exit;
      end;

      if not UseMRF then
        if frmMain.eOutputFile.Text = '' then
        begin
          ErrorMessage := ('Please select an output file.');
          lInputOk := False;
          Result := False;
          Exit;
        end;
    end;

    try
      if (Length(frmMain.omitcharbox.Text) > 0) and (frmMain.omitcharbox.Text[1] = chr(9)) then
        frmMain.omitcharbox.Text := chr(32);
    except
      ErrorMessage :=
        ('The omit character box is blank. Please specify an omit character code and retry the analysis.');
      FAllOk := False;
      lInputOk := False;
      Result := False;
      Exit;
    end;

    try
      if (Length(frmMain.nacharbox.Text) > 0) and (frmMain.nacharbox.Text[1] = chr(9)) then
        frmMain.nacharbox.Text := chr(32);
    except
      ErrorMessage :=
        ('The not administered character box is blank. Please specify a not administered character code and retry the analysis.');
      FAllOk := False;
      lInputOk := False;
      Result := False;
      Exit;
    end;

    if (Length(frmMain.omitcharbox.Text) > 0) and (Length(frmMain.nacharbox.Text) > 0) and
      (frmMain.omitcharbox.Text[1] = frmMain.nacharbox.Text[1]) then
    begin
      ErrorMessage := ('The omit character and not administered character are the same.' + #13 +
        'Please use different characters and retry the analysis.');
      FAllOk := False;
      lInputOk := False;
      Result := False;
      Exit;
    end; // check if omit and na use the same character (Feb.2013)

    if UseMRF and (frmMain.MRFText1 <> '') then
    begin
      mrfname := mrf1;
      AssignFile(lMRFFile, mrfname);
      ReWrite(lMRFFile);
      CloseFile(lMRFFile);
    end; // if reading in MRF

    if not(frmMain.cbDataMatrixFile.IsChecked) then
      if AItemResponsesBeginInColumn = 0 then
      begin
        ErrorMessage := ('The "Item responses begin in column" box is set to 0' + #13 +
          'Check the data file and specify a number greater than 0');
        FAllOk := False;
        lColumnOk := False;
        Result := False;
        Exit;
      end
      else if (AItemResponsesBeginInColumn <> 0) and
        ((StrToIntDef(frmMain.eTab2Numbers1.Text, 0) + StrToIntDef(frmMain.IDColumnsBegin.Text, 0)) >
        AItemResponsesBeginInColumn) then
      begin
        ErrorMessage := ('There are ' + frmMain.eTab2Numbers1.Text + ' columns of examinee ID that begin in column ' +
          frmMain.IDColumnsBegin.Text + #13 + 'The "item responses begin in column" box should be at least ' +
          IntToStr(StrToIntDef(frmMain.eTab2Numbers1.Text, 0) + StrToIntDef(frmMain.IDColumnsBegin.Text, 0)) +
          '. Please specify new value(s).');

        FAllOk := False;
        lColumnOk := False;
        Result := False;
        Exit;
      end;

    if lColumnOk and (frmMain.eTab2Numbers1.Text = '0') and not frmMain.cbDataMatrixFile.IsChecked then
    begin
      ErrorMessage := ('The number of examinee ID columns equals 0. We proceed with the analysis.');
      FAllOk := False;
      lColumnOk := False;
      Result := False;
      Exit;
    end;
  end;

  function CheckDataMatrixAndControlFiles(): Boolean;
  var
    ifor: Integer;
  begin
    Result := True;
    AssignFile(lDataMatrixFile, data);

    try
      Reset(lDataMatrixFile);
    except
      lInputOk := False;
      ErrorMessage := ('The data matrix file is already open in another program and cannot be opened by Iteman.' + #13 +
        'Please close the file, then hit Run againe to restart the analysis.');
      Result := False;
      CloseFile(lDataMatrixFile);
      Exit;
    end;

    if lInputOk then
    begin // checking to see if the data matrix has blank lines at the beginning
      Reset(lDataMatrixFile);
      Read(lDataMatrixFile, lChar);

      if (lChar = chr(13)) or (lChar = chr(26)) then
      begin
        lInputOk := False;
        ErrorMessage := ('The data matrix file has blank line(s) at the beginning.' + #13 +
          'Please delete the blank line(s) and try again.');
        Result := False;
        CloseFile(lDataMatrixFile);
        Exit;
      end;

      for ifor := 1 to 5 do
      begin
        Read(lDataMatrixFile, lChar);
        readformat[ifor] := lChar;
      end;

      isdocx := PWideChar(@readformat);

      if (Pos('docx', isdocx) <> 0) then
      begin
        lInputOk := False;
        ErrorMessage := ('The data matrix file is a MS WORD document.' + #13 +
          'Please open the document in a word processor and save as a Plain Text file.');

        CloseFile(lDataMatrixFile);
        Result := False;
        Exit;
      end;

      Reset(lDataMatrixFile);

      for ifor := 1 to 20 do
      begin
        Read(lDataMatrixFile, lChar);

        if lChar = chr(0) then
        begin
          lInputOk := False;
          ErrorMessage := ('The data matrix file is not an ANSI text file.' + #13 +
            'Please open the document in a word processor and save as a Plain Text file.');
          CloseFile(lDataMatrixFile);
          Exit;
        end;
      end;

      Reset(lDataMatrixFile);

      tabcounter := 0;
      commacounter := 0;

      while not eoln(lDataMatrixFile) do
      begin
        Read(lDataMatrixFile, lChar);

        if lChar = chr(9) then
          inc(tabcounter);
        if lChar = ',' then
          inc(commacounter);
      end;

      Reset(lDataMatrixFile);

      if frmMain.tabbox.IsChecked and (tabcounter = 0) then
      begin
        ErrorMessage := ('The data matrix file did not contain any tabs on the first line although the file was ' + #13
          + 'identified as tab delimited input. We proceed with the analysis without delimited input.');
        frmMain.delimitbox.IsChecked := False;
        frmMain.tabbox.IsChecked := False;
        lInputOk := True;
      end; // tab delim without tabs

      if frmMain.commabox.IsChecked and (commacounter = 0) then
      begin
        ErrorMessage := ('The data matrix file did not contain any commas on the first line although the file was ' +
          #13 + 'identified as comma delimited input. We proceed with the analysis without delimited input.');
        frmMain.delimitbox.IsChecked := False;
        frmMain.commabox.IsChecked := False;
        lInputOk := True;
      end; // comma delim without commas

      if (commacounter <> 0) and (tabcounter <> 0) then
      begin
        ErrorMessage := ('The data matrix file contained both tabs and commas for delimiters.' + #13 +
          'Proceed with the analysis using just tabs for the delimiters?' + #13 + #13 +
          'Note: Only click Yes if the commas were included as part of the examinee ID.');
        frmMain.delimitbox.IsChecked := True;
        frmMain.tabbox.IsChecked := True;
        frmMain.commabox.IsChecked := False;
        lInputOk := True;
      end; // comma and tab <> 0

      if (tabcounter <> 0) and not frmMain.tabbox.IsChecked then
      begin
        ErrorMessage := ('The data matrix file contains tabs.' + #13 +
          'We proceed with the analysis using tab delimited input');
        frmMain.delimitbox.IsChecked := True;
        frmMain.tabbox.IsChecked := True;
        lInputOk := True;
      end; // tabs in file without tab delim requested

      if (commacounter <> 0) and not frmMain.commabox.IsChecked then
      begin
        ErrorMessage := ('The data matrix file contains commas.' + #13 +
          'We change the program settings to use comma delimited input.');
        frmMain.delimitbox.IsChecked := True;
        frmMain.commabox.IsChecked := True;
      end; // commas in file without comma delim requested

      if not frmMain.cbDataMatrixFile.IsChecked then
      begin
        AssignFile(lItemInfoFile, ctrl);

        try
          Reset(lItemInfoFile);
        except
          lInputOk := False;
          ErrorMessage := ('The item control file is already open in another program and cannot be opened by Iteman.' +
            #13 + 'Please close the file, then hit Run again to restart the analysis.');
          Result := False;
          Exit;
        end;
      end; // if using FItem control file
    end;

    if lInputOk and not frmMain.cbDataMatrixFile.IsChecked then
    begin
      lInputOk := False;
      controldelim := chr(0);

      while not eoln(lItemInfoFile) do
      begin
        Read(lItemInfoFile, lChar);

        if lChar = chr(9) then
          lInputOk := True;
        if lChar = chr(9) then
          controldelim := chr(9);
      end;

      Reset(lItemInfoFile);

      if lInputOk = False then
        while not eoln(lItemInfoFile) do
        begin
          Read(lItemInfoFile, lChar);

          if lChar = ',' then
            lInputOk := True;
          if lChar = ',' then
            controldelim := ',';
        end;

      Reset(lItemInfoFile);

      if not lInputOk then
      begin
        ErrorMessage := ('The item control file is invalid.' + #13 +
          'The item control file must be a tab or comma delimited plain text file.');
        Result := False;
        Exit;
      end;
    end;
  end;

  function CheckControlAndDataItemsCount(): Boolean;
  var
    lNumMatch: Boolean;
    lRow: Integer;
    lColumn: Integer;
  begin
    Result := True;

    if lInputOk and frmMain.cbDataMatrixFile.IsChecked then
    begin
      Reset(lDataMatrixFile);

      lNumMatch := True;
      try
        Read(lDataMatrixFile, FNumberOfItems);
      except
        FAllOk := False;
      end;

      Read(lDataMatrixFile, lChar);
      Read(lDataMatrixFile, lChar);
      Read(lDataMatrixFile, lChar);
      Read(lDataMatrixFile, lChar);

      try
        Read(lDataMatrixFile, lExamineeIDColumns);
      except
        FAllOk := False;
      end;

      ReadLn(lDataMatrixFile);

      for lRow := 1 to 3 do
        ReadLn(lDataMatrixFile);

      lNumberOfEntries := 0;

      while not eoln(lDataMatrixFile) do
      begin
        Read(lDataMatrixFile, lChar);
        inc(lNumberOfEntries);
      end;

      if frmMain.difcheckbox.IsChecked and (lNumberOfEntries <> (FNumberOfItems + succ(lExamineeIDColumns))) then
        lNumMatch := False;

      if not frmMain.difcheckbox.IsChecked and (lNumberOfEntries <> (FNumberOfItems + lExamineeIDColumns)) then
        lNumMatch := False;

      if not lNumMatch then
      begin
        ErrorMessage := ('The number of items in your Item Control File and the first line of your data file mismatch.'
          + #13 + 'Please correct this error and re-run your data. This run is terminated.');
        lInputOk := False;
        Result := False;
        Exit;
      end;
    end;

    if not frmMain.cbDataMatrixFile.IsChecked then
    begin
      FNumberOfItems := 0;

      while not EOF(lItemInfoFile) do
      begin
        ReadLn(lItemInfoFile, lChar);
        // chr(13) is return, so blank lines are not treated as items

        if lChar <> chr(13) then
          inc(FNumberOfItems);
      end;

      Reset(lItemInfoFile);

      if (GetColumnsCountOfCSVFile(data) - DataMatrixAdditionalColumns - ADataMatrixVariableAdditionalColumns) <>
        (FNumberOfItems - AItemControlVariableAdditionalRows) then
      begin
        ErrorMessage := ('The number of items in your Item Control File and the first line of your data file mismatch.'
          + #13 + 'Please correct this error and re-run your data.  This run is terminated.');
        lInputOk := False;
        Result := False;
        Exit;
      end;

      Reset(lItemInfoFile);
      Reset(lDataMatrixFile);
    end;
  end;

  function CheckOutputAndScoreFile(): Boolean;
  var
    lScoreFileName: string;
  begin
    Result := True;

    if lInputOk then
    begin
      if UseMRF then
        lCSVFileName := ChangeFileExt(outp, ' Output.csv')
      else
        lCSVFileName := ChangeFileExt(outp, '.csv');

      AssignFile(lOutputFile, lCSVFileName);

      try
        ReWrite(lOutputFile);
      except
        lInputOk := False;
        ErrorMessage := ('The Main CSV output file is already open in another program and cannot be saved by Iteman.' +
          #13 + 'Please close the file, then hit Run to restart the analysis.');
        Result := False;
        Exit;
      end;

      lScoreFileName := ChangeFileExt(outp, '') + ' Scores';
      lScoreFileName := ChangeFileExt(lScoreFileName, '.csv');
      AssignFile(lScoresFile, lScoreFileName);

      try
        ReWrite(lScoresFile);
      except
        lInputOk := False;
        ErrorMessage := ('The Scores CSV file is already open in another program and cannot be saved by Iteman.' + #13 +
          'Please close the file, then hit Run to restart the analysis.');
        Result := False;
        Exit;
      end;

      if frmMain.ScoredMatrixBox.IsChecked then
      begin
        lCSVFileName := ChangeFileExt(ChangeFileExt(outp, '') + ' Matrix', '.txt');
        AssignFile(lScoredMatrixFile, lCSVFileName);

        try
          ReWrite(lScoredMatrixFile);
        except
          lInputOk := False;
          ErrorMessage :=
            ('The scored data matrix file is already open in another program and cannot be saved by Iteman.' + #13 +
            'Please close the file, then hit Retry to continue the analysis.');
          Result := False;
          Exit;
        end;
      end;

      if frmMain.cbExternalScoreFile.IsChecked then
      begin
        AssignFile(lExternalScoreFile, frmMain.eExternalScoreFile.Text);

        try
          Reset(lExternalScoreFile);
        except
          lInputOk := False;
          ErrorMessage := ('The external score file is already open in another program and cannot be opened by Iteman.'
            + #13 + 'Please close the file, then hit Retry to continue the analysis.');
          Result := False;
          Exit;
        end;
      end;
    end;
  end;

  procedure CSVOutputSection();
  var
    lColumn: Integer;
    lRow: Integer;
    lDomain: Integer;
    dfor: Integer;
    kfor, jfor: Integer;
  begin
    frmMain.ProgressBar.Value := frmMain.ProgressBar.Value + 10;
    frmMain.ProgressLabel.Text := 'Producing CSV Output';
    Application.ProcessMessages;

    if (StrToIntDef(frmMain.Precision.Text, 0) = 2) then
      prec := '0.00';
    if (StrToIntDef(frmMain.Precision.Text, 0) = 3) then
      prec := '0.000';
    if (StrToIntDef(frmMain.Precision.Text, 0) = 4) then
      prec := '0.0000';

    { Write examinee results to scores csv file }
    if not scoredrate then
      Write(lScoresFile,
        'Sequence,ID,All Items,Scored Items,Pretest items,Scored Proportion Correct,Rank,Percentile,Group')
    else
      Write(lScoresFile, 'Sequence,ID,All Items,Scored Items,Pretest items,Rank,Percentile,Group');

    if frmMain.cbExternalScoreFile.IsChecked then
      Write(lScoresFile, ',External');
    if FNumDomains > 1 then
      For lDomain := 1 to FNumDomains do
        write(lScoresFile, ',', FStoreDomain[lDomain]);
    if frmMain.ScaledCheckBox.IsChecked then
      Write(lScoresFile, ',Scaled Total Score');
    if (FNumDomains > 1) and (frmMain.ScaledDomain.IsChecked) then
      for lDomain := 1 to FNumDomains do
        write(lScoresFile, ',Scaled ', FStoreDomain[lDomain]);

    if frmMain.classcheckbox.IsChecked then
      write(lScoresFile, ',Classification');
    if not scoredrate then
      write(lScoresFile, ',CSEM III,CSEM IV');

    writeln(lScoresFile);
    FProppassing := 0;

    for lRow := 1 to FNumberOfExaminees do
    begin
      Write(lScoresFile, lRow, ',', lExamineeIDArray[lRow], ',', lScoreArray[lRow, 4], ',', lScoreArray[lRow, 1], ',',
        lScoreArray[lRow, 5]);

      if not scoredrate then
        write(lScoresFile, ',', formatfloat(prec, (lScoreArray[lRow, 1] / FTestStatsArray[1, 1])));

      write(lScoresFile, ',', lRankingArray[lRow, 1]);
      percentile := FNumberOfExaminees - (lRankingArray[lRow, 1]);
      write(lScoresFile, ',', formatfloat('0.00', ((percentile) / FNumberOfExaminees) * 100), '%');
      write(lScoresFile, ',', lRankingArray[lRow, 2]);

      if frmMain.cbExternalScoreFile.IsChecked then
        Write(lScoresFile, ',', FExternalScoreArray[lRow]:2 + pr:pr);
      if FNumDomains > 1 then
        for lDomain := 1 to FNumDomains do
          write(lScoresFile, ',', FDomainArray[lRow, lDomain]);

      if frmMain.ScaledCheckBox.IsChecked then
        Write(lScoresFile, ',', FScaledScoreArray[lRow]:2 + pr:pr);

      if (FNumDomains > 1) and (frmMain.ScaledDomain.IsChecked) then
        for lDomain := 1 to FNumDomains do
          Write(lScoresFile, ',', lDomScaleArray[lRow, lDomain]:2 + pr:pr);

      if frmMain.classcheckbox.IsChecked then
      begin
        if frmMain.ScoredClass.IsChecked then
        begin
          if (lScoreArray[lRow, 1] >= StrToFloatDef(frmMain.classcutbox.Text, 0)) then
          begin
            write(lScoresFile, ',', frmMain.HighLabel.Text);
            FProppassing := FProppassing + 1;
          end
          else
            write(lScoresFile, ',', frmMain.LowLabel.Text);
        end
        else if frmMain.ScaledClass.IsChecked then
        begin
          if (FScaledScoreArray[lRow] >= StrToFloatDef(frmMain.classcutbox.Text, 0)) then
          begin
            write(lScoresFile, ',', frmMain.HighLabel.Text);
            FProppassing := FProppassing + 1;
          end
          else
            write(lScoresFile, ',', frmMain.LowLabel.Text);
        end;

        // old
        // if frmMain.ScoredClass.IsChecked and (lScoreArray[lRow,1] >= StrToFloatDef(frmMain.classcutbox.Text, 0)) then
        // write(lScoresFile, ',', frmMain.HighLabel.Text)
        // else
        // write(lScoresFile, ',', frmMain.LowLabel.Text);
        //
        // if frmMain.ScoredClass.IsChecked and (lScoreArray[lRow,1] >= StrToFloatDef(frmMain.classcutbox.Text, 0)) then
        // FProppassing:= FProppassing + 1;
        //
        // if frmMain.ScaledClass.IsChecked and (FScaledScoreArray[lRow] >= StrToFloatDef(frmMain.classcutbox.Text, 0)) then
        // write(lScoresFile,',', frmMain.HighLabel.Text)
        // else
        // write(lScoresFile,',',frmMain.LowLabel.Text);
        //
        // if frmMain.ScaledClass.IsChecked and (FScaledScoreArray[lRow] >= StrToFloatDef(frmMain.classcutbox.Text, 0)) then
        // FProppassing:= FProppassing + 1;

      end;

      if not scoredrate then
      begin
        lRealTest := lCsemArray[lScoreArray[lRow, 1], 1];
        write(lScoresFile, ',', formatfloat(prec, sqrt(abs(lRealTest))));
        write(lScoresFile, ',', formatfloat(prec, sqrt(abs(lCsemArray[lScoreArray[lRow, 1], 2]))));
      end;

      writeln(lScoresFile);
    end;

    FProppassing := FProppassing / FNumberOfExaminees;

    { Write overall stats table to output csv file }
    writeln(lOutputFile, 'Test Information');
    writeln(lOutputFile, 'Run title:,', lRunTitle);
    writeln(lOutputFile, 'Date:,', DateTimeToStr(Date));
    writeln(lOutputFile, 'N:,', IntToStr(FNumberOfExaminees));
    writeln(lOutputFile, 'Total items:,', IntToStr(FNumberOfItems));
    writeln(lOutputFile, 'Scored items:,', IntToStr(FScoredItems));
    writeln(lOutputFile, 'Pretest items:,', IntToStr(round(FTestStatsArray[3, 1])));

    if FNumDomains > 1 then
      writeln(lOutputFile, 'Domains:,', IntToStr(FNumDomains));

    writeln(lOutputFile, 'Minimum P:,', FMinP:3 + pr:pr);
    writeln(lOutputFile, 'Maximum P:,', FMaxP:3 + pr:pr);
    writeln(lOutputFile, 'Minimum Item Mean:,', StrToFloatDef(frmMain.MinMeanBox.Text, 0):3 + pr:pr);
    writeln(lOutputFile, 'Maximum Item Mean:,', StrToFloatDef(frmMain.MaxMeanBox.Text, 0):3 + pr:pr);
    writeln(lOutputFile, 'Minimum Item R:,', FMinR:3 + pr:pr);
    writeln(lOutputFile, 'Maximum Item R:,', FMaxR:3 + pr:pr);
    writeln(lOutputFile);
    writeln(lOutputFile);
    writeln(lOutputFile, 'Summary statistics');
    writeln(lOutputFile);

    { FTestStatsArray:  NItems, Mean, SD, Min, Max, Alpha, Mean FP, Mean Rpbis For Scored, Total, and FPreTest }

    { Write overall stats table to output csv file }
    if FPreTest > 0 then
      Write(lOutputFile, 'Statistic,Scored Items,All Items,Pretest Items')
    else
      Write(lOutputFile, 'Statistic,Scored Items');

    if frmMain.ScaledCheckBox.IsChecked then
      Write(lOutputFile, ',Scaled');
    if frmMain.cbExternalScoreFile.IsChecked then
      Write(lOutputFile, ',External');
    if FNumDomains > 1 then
      for lDomain := 1 to FNumDomains do
        write(lOutputFile, ',', FStoreDomain[lDomain]);

    writeln(lOutputFile);

    if FPreTest > 0 then
      write(lOutputFile, 'Items:,', FTestStatsArray[1, 1]:3 + pr:pr, ',', FTestStatsArray[2, 1]:3 + pr:pr, ',',
        FTestStatsArray[3, 1]:3 + pr:pr)
    else
      write(lOutputFile, 'Items:,', FTestStatsArray[1, 1]:3 + pr:pr);

    if frmMain.ScaledCheckBox.IsChecked then
      Write(lOutputFile, ',', FTestStatsArray[4, 1]:3 + pr:pr);
    if frmMain.cbExternalScoreFile.IsChecked then
      Write(lOutputFile, ',', FTestStatsArray[5, 1]:3 + pr:pr);
    if FNumDomains > 1 then
      for lDomain := 1 to FNumDomains do
        write(lOutputFile, ',', FDomainStatsArray[lDomain, 1]:3 + pr:pr);

    writeln(lOutputFile);

    if FPreTest > 0 then
      write(lOutputFile, 'Mean:,', FTestStatsArray[1, 2]:3 + pr:pr, ',', FTestStatsArray[2, 2]:3 + pr:pr, ',',
        FTestStatsArray[3, 2]:3 + pr:pr)
    else
      write(lOutputFile, 'Mean:,', FTestStatsArray[1, 2]:3 + pr:pr);

    if frmMain.ScaledCheckBox.IsChecked then
      Write(lOutputFile, ',', FTestStatsArray[4, 2]:3 + pr:pr);
    if frmMain.cbExternalScoreFile.IsChecked then
      Write(lOutputFile, ',', FTestStatsArray[5, 2]:3 + pr:pr);
    if FNumDomains > 1 then
      for lDomain := 1 to FNumDomains do
        write(lOutputFile, ',', FDomainStatsArray[lDomain, 2]:3 + pr:pr);

    writeln(lOutputFile);

    if FPreTest > 0 then
      write(lOutputFile, 'SD:,', FTestStatsArray[1, 3]:3 + pr:pr, ',', FTestStatsArray[2, 3]:3 + pr:pr, ',',
        FTestStatsArray[3, 3]:3 + pr:pr)
    else
      write(lOutputFile, 'SD:,', FTestStatsArray[1, 3]:3 + pr:pr);

    if frmMain.ScaledCheckBox.IsChecked then
      Write(lOutputFile, ',', FTestStatsArray[4, 3]:3 + pr:pr);
    if frmMain.cbExternalScoreFile.IsChecked then
      Write(lOutputFile, ',', FTestStatsArray[5, 3]:3 + pr:pr);
    if FNumDomains > 1 then
      for lDomain := 1 to FNumDomains do
        write(lOutputFile, ',', FDomainStatsArray[lDomain, 3]:3 + pr:pr);

    writeln(lOutputFile);

    if FPreTest > 0 then
      write(lOutputFile, 'Min Score:,', FTestStatsArray[1, 4]:3 + pr:pr, ',', FTestStatsArray[2, 4]:3 + pr:pr, ',',
        FTestStatsArray[3, 4]:3 + pr:pr)
    else
      write(lOutputFile, 'Min Score:,', FTestStatsArray[1, 4]:3 + pr:pr);

    if frmMain.ScaledCheckBox.IsChecked then
      Write(lOutputFile, ',', FTestStatsArray[4, 4]:3 + pr:pr);
    if frmMain.cbExternalScoreFile.IsChecked then
      Write(lOutputFile, ',', FTestStatsArray[5, 4]:3 + pr:pr);
    if FNumDomains > 1 then
      for lDomain := 1 to FNumDomains do
        write(lOutputFile, ',', FDomainStatsArray[lDomain, 4]:3 + pr:pr);

    writeln(lOutputFile);

    if FPreTest > 0 then
      write(lOutputFile, 'Max Score:,', FTestStatsArray[1, 5]:3 + pr:pr, ',', FTestStatsArray[2, 5]:3 + pr:pr, ',',
        FTestStatsArray[3, 5]:3 + pr:pr)
    else
      write(lOutputFile, 'Max Score:,', FTestStatsArray[1, 5]:3 + pr:pr);

    if frmMain.ScaledCheckBox.IsChecked then
      Write(lOutputFile, ',', FTestStatsArray[4, 5]:3 + pr:pr);
    if frmMain.cbExternalScoreFile.IsChecked then
      Write(lOutputFile, ',', FTestStatsArray[5, 5]:3 + pr:pr);

    if FNumDomains > 1 then
      for lDomain := 1 to FNumDomains do
        write(lOutputFile, ',', FDomainStatsArray[lDomain, 5]:3 + pr:pr);

    writeln(lOutputFile);

    if FIsMult then
    begin
      if FPreTest > 0 then
        write(lOutputFile, 'Mean P:,', FTestStatsArray[1, 7]:3 + pr:pr, ',', FTestStatsArray[2, 7]:3 + pr:pr, ',',
          FTestStatsArray[3, 7]:3 + pr:pr)
      else
        write(lOutputFile, 'Mean P:,', FTestStatsArray[1, 7]:3 + pr:pr);

      if frmMain.ScaledCheckBox.IsChecked then
        Write(lOutputFile, ',-');
      if frmMain.cbExternalScoreFile.IsChecked then
        Write(lOutputFile, ',-');
      if FNumDomains > 1 then
        for lDomain := 1 to FNumDomains do
          write(lOutputFile, ',', FDomainStatsArray[lDomain, 7]:3 + pr:pr);

      writeln(lOutputFile);
    end; // if FIsMult = True

    if FIsRate then
    begin
      if FPreTest > 0 then
        write(lOutputFile, 'Item Mean:,', FTestStatsArray[1, 16]:3 + pr:pr, ',', FTestStatsArray[2, 16]:3 + pr:pr, ',',
          FTestStatsArray[3, 16]:3 + pr:pr)
      else
        write(lOutputFile, 'Item Mean:,', FTestStatsArray[1, 16]:3 + pr:pr);

      if frmMain.ScaledCheckBox.IsChecked then
        Write(lOutputFile, ',-');
      if frmMain.cbExternalScoreFile.IsChecked then
        Write(lOutputFile, ',-');
      if FNumDomains > 1 then
        for lDomain := 1 to FNumDomains do
          write(lOutputFile, ',', FDomainStatsArray[lDomain, 16]:3 + pr:pr);

      writeln(lOutputFile);
    end;

    if FPreTest > 0 then
      write(lOutputFile, 'Mean R:,', FTestStatsArray[1, 8]:3 + pr:pr, ',', FTestStatsArray[2, 8]:3 + pr:pr, ',',
        FTestStatsArray[3, 8]:3 + pr:pr)
    else
      write(lOutputFile, 'Mean R:,', FTestStatsArray[1, 8]:3 + pr:pr);

    if frmMain.ScaledCheckBox.IsChecked then
      Write(lOutputFile, ',-');
    if frmMain.cbExternalScoreFile.IsChecked then
      Write(lOutputFile, ',-');
    if FNumDomains > 1 then
      for lDomain := 1 to FNumDomains do
        write(lOutputFile, ',', FDomainStatsArray[lDomain, 8]:3 + pr:pr);

    writeln(lOutputFile);
    writeln(lOutputFile);

    writeln(lOutputFile, 'Reliability Analysis');
    writeln(lOutputFile);
    { ----Headers for Reliability Analysis!---- }
    if FPreTest > 0 then
      Write(lOutputFile, 'Statistic,Scored Items,All Items,Pretest Items')
    else
      Write(lOutputFile, 'Statistic,Scored Items');

    if frmMain.ScaledCheckBox.IsChecked then
      Write(lOutputFile, ',Scaled');
    if frmMain.cbExternalScoreFile.IsChecked then
      Write(lOutputFile, ',External');
    if FNumDomains > 1 then
      for lDomain := 1 to FNumDomains do
        write(lOutputFile, ',', FStoreDomain[lDomain]);

    writeln(lOutputFile);

    if FPreTest > 0 then
      write(lOutputFile, 'Alpha:,', FTestStatsArray[1, 6]:3 + pr:pr, ',', FTestStatsArray[2, 6]:3 + pr:pr, ',',
        FTestStatsArray[3, 6]:3 + pr:pr)
    else
      write(lOutputFile, 'Alpha:,', FTestStatsArray[1, 6]:3 + pr:pr);

    if frmMain.ScaledCheckBox.IsChecked then
      Write(lOutputFile, ',-');
    if frmMain.cbExternalScoreFile.IsChecked then
      Write(lOutputFile, ',-');
    if FNumDomains > 1 then
      for lDomain := 1 to FNumDomains do
        write(lOutputFile, ',', FDomainStatsArray[lDomain, 6]:3 + pr:pr);

    writeln(lOutputFile);

    if FPreTest > 0 then
      write(lOutputFile, 'SEM:,', FTestStatsArray[1, 9]:3 + pr:pr, ',', FTestStatsArray[2, 9]:3 + pr:pr, ',',
        FTestStatsArray[3, 9]:3 + pr:pr)
    else
      write(lOutputFile, 'SEM:,', FTestStatsArray[1, 9]:3 + pr:pr);

    if frmMain.ScaledCheckBox.IsChecked then
      Write(lOutputFile, ',-');
    if frmMain.cbExternalScoreFile.IsChecked then
      Write(lOutputFile, ',-');
    if FNumDomains > 1 then
      for lDomain := 1 to FNumDomains do
        write(lOutputFile, ',', FDomainStatsArray[lDomain, 9]:3 + pr:pr);

    writeln(lOutputFile);

    if FPreTest > 0 then
      write(lOutputFile, 'Split-Half (Random):,', FTestStatsArray[1, 10]:3 + pr:pr, ',', FTestStatsArray[2, 10]:3 + pr
        :pr, ',', '-')
    else
      write(lOutputFile, 'Split-Half (Random):,', FTestStatsArray[2, 10]:3 + pr:pr);

    if frmMain.ScaledCheckBox.IsChecked then
      Write(lOutputFile, ',-');
    if frmMain.cbExternalScoreFile.IsChecked then
      Write(lOutputFile, ',-');
    if FNumDomains > 1 then
      for lDomain := 1 to FNumDomains do
        write(lOutputFile, ',', FDomainStatsArray[lDomain, 10]:3 + pr:pr);

    writeln(lOutputFile);

    if FPreTest > 0 then
      write(lOutputFile, 'Split-Half (First-Last):,', FTestStatsArray[1, 11]:3 + pr:pr, ',', FTestStatsArray[2, 11]
        :3 + pr:pr, ',', '-')
    else
      write(lOutputFile, 'Split-Half (First-Last):,', FTestStatsArray[2, 11]:3 + pr:pr);

    if frmMain.ScaledCheckBox.IsChecked then
      Write(lOutputFile, ',-');
    if frmMain.cbExternalScoreFile.IsChecked then
      Write(lOutputFile, ',-');
    if FNumDomains > 1 then
      for lDomain := 1 to FNumDomains do
        write(lOutputFile, ',', FDomainStatsArray[lDomain, 11]:3 + pr:pr);

    writeln(lOutputFile);

    if FPreTest > 0 then
      write(lOutputFile, 'Split-Half (Odd-Even):,', FTestStatsArray[1, 12]:3 + pr:pr, ',', FTestStatsArray[2, 12]:3 + pr
        :pr, ',', '-')
    else
      write(lOutputFile, 'Split-Half (Odd-Even):,', FTestStatsArray[2, 12]:3 + pr:pr);

    if frmMain.ScaledCheckBox.IsChecked then
      Write(lOutputFile, ',-');
    if frmMain.cbExternalScoreFile.IsChecked then
      Write(lOutputFile, ',-');
    if FNumDomains > 1 then
      for lDomain := 1 to FNumDomains do
        write(lOutputFile, ',', FDomainStatsArray[lDomain, 12]:3 + pr:pr);

    writeln(lOutputFile);

    if FPreTest > 0 then
      write(lOutputFile, 'S-B Random:,', FTestStatsArray[1, 13]:3 + pr:pr, ',', FTestStatsArray[2, 13]:3 + pr
        :pr, ',', '-')
    else
      write(lOutputFile, 'S-B Random:,', FTestStatsArray[2, 13]:3 + pr:pr);

    if frmMain.ScaledCheckBox.IsChecked then
      Write(lOutputFile, ',-');
    if frmMain.cbExternalScoreFile.IsChecked then
      Write(lOutputFile, ',-');
    if FNumDomains > 1 then
      for lDomain := 1 to FNumDomains do
        write(lOutputFile, ',', FDomainStatsArray[lDomain, 13]:3 + pr:pr);

    writeln(lOutputFile);

    if FPreTest > 0 then
      write(lOutputFile, 'S-B First-Last:,', FTestStatsArray[1, 14]:3 + pr:pr, ',', FTestStatsArray[2, 14]:3 + pr
        :pr, ',', '-')
    else
      write(lOutputFile, 'S-B First-Last:,', FTestStatsArray[2, 14]:3 + pr:pr);

    if frmMain.ScaledCheckBox.IsChecked then
      Write(lOutputFile, ',-');
    if frmMain.cbExternalScoreFile.IsChecked then
      Write(lOutputFile, ',-');
    if FNumDomains > 1 then
      for lDomain := 1 to FNumDomains do
        write(lOutputFile, ',', FDomainStatsArray[lDomain, 14]:3 + pr:pr);

    writeln(lOutputFile);

    if FPreTest > 0 then
      write(lOutputFile, 'S-B Odd-Even:,', FTestStatsArray[1, 15]:3 + pr:pr, ',', FTestStatsArray[2, 15]:3 + pr
        :pr, ',', '-')
    else
      write(lOutputFile, 'S-B Odd-Even:,', FTestStatsArray[2, 15]:3 + pr:pr);

    if frmMain.ScaledCheckBox.IsChecked then
      Write(lOutputFile, ',-');
    if frmMain.cbExternalScoreFile.IsChecked then
      Write(lOutputFile, ',-');
    if FNumDomains > 1 then
      for lDomain := 1 to FNumDomains do
        write(lOutputFile, ',', FDomainStatsArray[lDomain, 15]:3 + pr:pr);

    writeln(lOutputFile);

    if FNumDomains > 1 then // subscore intercorrelation added Mar.2013
    begin
      writeln(lOutputFile);
      writeln(lOutputFile, 'Subscore Intercorrelation');
      writeln(lOutputFile);
      write(lOutputFile, 'Domain');

      for lDomain := 1 to FNumDomains do
        write(lOutputFile, ',', FStoreDomain[lDomain]);

      writeln(lOutputFile);

      for lRow := 1 to FNumDomains do
      begin
        write(lOutputFile, FStoreDomain[lRow]);

        for kfor := 1 to FNumDomains do
          write(lOutputFile, ',', FDomainCorrArray[lRow, kfor]:3 + pr:pr);

        writeln(lOutputFile);
      end;
    end;

    writeln(lOutputFile);
    writeln(lOutputFile, 'Item statistics - row format');
    writeln(lOutputFile);

    { Write FItem stat headers to output csv file }
    if FIsMult and not FIsRate then
      if FNumDomains > 1 then
        Write(lOutputFile,
          'Sequence,Item ID,Key,Scored,NumOptions,Domain,N,P,Total Rpbis,Total Rbis,Domain Rpbis,Domain Rbis,Alpha w/o,Flags,Omit Freq,NR Freq,')
      else
        Write(lOutputFile,
          'Sequence,Item ID,Key,Scored,NumOptions,Domain,N,P,Total Rpbis,Total Rbis,Alpha w/o,Flags,Omit Freq,NR Freq,');

    if FIsMult and FIsRate then
      if FNumDomains > 1 then
        Write(lOutputFile,
          'Sequence,Item ID,Key,Scored,NumOptions,Domain,N,Mean P/Item Mean,Total Rpbis/R,Total Rbis/R,Domain Rpbis/R,Domain Rbis/R,Alpha w/o,Flags,Omit Freq,NR Freq,')
      else
        Write(lOutputFile,
          'Sequence,Item ID,Key,Scored,NumOptions,Domain,N,Mean P/Item Mean,Total Rpbis/R,Total Rbis/R,Alpha w/o,Flags,Omit Freq,NR Freq,');

    if not FIsMult and FIsRate then
      if FNumDomains > 1 then
        Write(lOutputFile,
          'Sequence,Item ID,Key,Scored,NumOptions,Domain,N,Item Mean,Total R,Total Eta,Domain R,Domain Eta,Alpha w/o,Flags,Omit Freq,NR Freq,')
      else
        Write(lOutputFile,
          'Sequence,Item ID,Key,Scored,NumOptions,Domain,N,Item Mean,Total R,Total Eta,Alpha w/o,Flags,Omit Freq,NR Freq,');

    if lAllPC then
      pcint := 1
    else
      pcint := 0;

    for lColumn := 1 to FMaxItemOptions do
      Write(lOutputFile, lColumn - pcint, ' Freq,');

    if not frmMain.OmitAsinc.IsChecked then
      write(lOutputFile, 'Omit Prop,');

    for lColumn := 1 to FMaxItemOptions do
      Write(lOutputFile, lColumn - pcint, ' Prop,');

    if not frmMain.OmitAsinc.IsChecked then
      write(lOutputFile, 'Omit Rpbis,');

    for lColumn := 1 to FMaxItemOptions do
      Write(lOutputFile, lColumn - pcint, ' Rpbis,');

    if not frmMain.OmitAsinc.IsChecked then
      write(lOutputFile, 'Omit Rbis,');

    for lColumn := 1 to FMaxItemOptions do
      Write(lOutputFile, lColumn - pcint, ' Rbis,');

    if not frmMain.OmitAsinc.IsChecked then
      write(lOutputFile, 'Omit Mean,');

    for lColumn := 1 to FMaxItemOptions do
      Write(lOutputFile, lColumn - pcint, ' Mean,');

    if not frmMain.OmitAsinc.IsChecked then
      write(lOutputFile, 'Omit SD,');

    for lColumn := 1 to FMaxItemOptions do
      Write(lOutputFile, lColumn - pcint, ' SD,');

    for lColumn := 1 to FMaxItemOptions do
      for jfor := 1 to FNumCut do
        Write(lOutputFile, lColumn - pcint, ' Grp', IntToStr(jfor), ' P,');

    writeln(lOutputFile);

    { Write Item stats to output csv file, Row style }
    for lRow := 1 to FNumberOfItems do
      if lItemInfoArray[lRow, 2] <> 'N' then
      begin
        FFlagID[lRow] := 0; // resetting to zero

        { This is columns pertaining to FItem info }
        if lItemInfoArray[lRow, 2] = 'P' then
          PreText := 'Pretest'
        else
          PreText := 'Yes';

        if lItemInfoArray[lRow, 2] = 'Y' then
          PreStr := formatfloat('0.000', FItemStatsArray[lRow, 5])
        else
          PreStr := 'NA';

        if lNKeyArray[lRow, 1] = 1 then
          Write(lOutputFile, lRow, ',', lItemIDArray[lRow], ',', lItemInfoArray[lRow, 1], ',', PreText, ',',
            lItemScaleArray[lRow, 1], ',', FStoreDomain[lItemScaleArray[lRow, 2]], ',', FItemStatsArray[lRow, 1]:1 + pr
            :pr, ',', FItemStatsArray[lRow, 2]:1 + pr:pr, ',', FItemStatsArray[lRow, 3]:1 + pr:pr, ',',
            FItemStatsArray[lRow, 4]:1 + pr:pr, ',');

        if lNKeyArray[lRow, 1] > 1 then
        begin
          Write(lOutputFile, lRow, ',', lItemIDArray[lRow], ',', lItemInfoArray[lRow, 1]);

          for dfor := 1 to lNKeyArray[lRow, 1] - 1 do
            write(lOutputFile, ' ', UpperCase(lItemInfoArray[lRow, 3 + dfor]));

          write(lOutputFile, ',', PreText, ',', lItemScaleArray[lRow, 1], ',', FStoreDomain[lItemScaleArray[lRow, 2]],
            ',', FItemStatsArray[lRow, 1]:1 + pr:pr, ',', FItemStatsArray[lRow, 2]:1 + pr:pr, ',',
            FItemStatsArray[lRow, 3]:1 + pr:pr, ',', FItemStatsArray[lRow, 4]:1 + pr:pr, ',');
        end;

        if FNumDomains > 1 then
          Write(lOutputFile, FItemStatsArray[lRow, 7]:1 + pr:pr, ',', FItemStatsArray[lRow, 8]:1 + pr:pr, ',');

        write(lOutputFile, PreStr, ','); // Alpha without

        lCheckKey := False;
        lItemIsFlagged := False;
        { Flags }
        for jfor := 1 to lItemScaleArray[lRow, 1] do
          if (lItemInfoArray[lRow, 3] = 'M') and (FItemStatsArray[lRow, 6] <> 0) then
            if FItemStatsArray[lRow, 3] < FOptionRpbisArray[lRow, jfor] then
              lCheckKey := True;

        if lCheckKey then
        begin
          write(lOutputFile, FFlagArray[1], ' ');
          lItemIsFlagged := True;
        end;

        if lItemInfoArray[lRow, 3] = 'M' then
        begin
          if FItemStatsArray[lRow, 2] < FMinP then
            write(lOutputFile, FFlagArray[2], ' ');
          if FItemStatsArray[lRow, 2] < FMinP then
            lItemIsFlagged := True;
          if FItemStatsArray[lRow, 2] > FMaxP then
            write(lOutputFile, FFlagArray[3], ' ');
          if FItemStatsArray[lRow, 2] > FMaxP then
            lItemIsFlagged := True;
        end;

        if lItemInfoArray[lRow, 3] <> 'M' then
        begin
          if FItemStatsArray[lRow, 2] < FMinMean then
            write(lOutputFile, FFlagArray[6], ' ');
          if FItemStatsArray[lRow, 2] < FMinMean then
            lItemIsFlagged := True;
          if FItemStatsArray[lRow, 2] > FMaxMean then
            write(lOutputFile, FFlagArray[7], ' ');
          if FItemStatsArray[lRow, 2] > FMaxMean then
            lItemIsFlagged := True;
        end;

        if FItemStatsArray[lRow, 3] < FMinR then
          write(lOutputFile, FFlagArray[4], ' ');
        if FItemStatsArray[lRow, 3] < FMinR then
          lItemIsFlagged := True;
        if FItemStatsArray[lRow, 3] > FMaxR then
          write(lOutputFile, FFlagArray[5], ' ');
        if FItemStatsArray[lRow, 3] > FMaxR then
          lItemIsFlagged := True;

        if (lItemInfoArray[lRow, 9] = 'D') and (frmMain.difcheckbox.IsChecked) then
        begin
          if lDifarray[lRow, 5] < 0.05 then
            lItemIsFlagged := True;
          if lDifarray[lRow, 5] < 0.05 then
            write(lOutputFile, FFlagArray[8], ' ');
        end; // DIF coding

        if lItemIsFlagged then
          inc(FNumFlagItems);
        if lItemIsFlagged then
          FFlagID[lRow] := lRow;

        write(lOutputFile, ',', IntToStr(FOptionCountsArray[lRow, 16]), ',', IntToStr(FOptionCountsArray[lRow, 17]));

        { This is columns of option counts }
        for lColumn := 1 to FMaxItemOptions do
          if lColumn < lItemScaleArray[lRow, 1] + 1 then
            Write(lOutputFile, ',', FOptionCountsArray[lRow, lColumn])
          else
            Write(lOutputFile, ',');

        { This is columns of option proportions }
        if not frmMain.OmitAsinc.IsChecked then
          write(lOutputFile, ',', FOptionPropsArray[lRow, 16]:1 + pr:pr);

        for lColumn := 1 to FMaxItemOptions do
          if lColumn < lItemScaleArray[lRow, 1] + 1 then
            Write(lOutputFile, ',', FOptionPropsArray[lRow, lColumn]:1 + pr:pr)
          else
            Write(lOutputFile, ',');

        { This is columns of option Rpbis }
        if not frmMain.OmitAsinc.IsChecked then
          write(lOutputFile, ',', FOptionRpbisArray[lRow, 16]:1 + pr:pr);

        for lColumn := 1 to FMaxItemOptions do
          if FItemStatsArray[lRow, 2] > 0 then
          begin
            if lColumn < lItemScaleArray[lRow, 1] + 1 then
              Write(lOutputFile, ',', FOptionRpbisArray[lRow, lColumn]:1 + pr:pr)
            else
              Write(lOutputFile, ',');
          end
          else
            Write(lOutputFile, '-,');

        { This is columns of option Rbis }
        if not frmMain.OmitAsinc.IsChecked then
          write(lOutputFile, ',', FOptionRbisArray[lRow, 16]:1 + pr:pr);

        for lColumn := 1 to FMaxItemOptions do
          if FItemStatsArray[lRow, 2] > 0 then
          begin
            if lColumn < lItemScaleArray[lRow, 1] + 1 then
              Write(lOutputFile, ',', FOptionRbisArray[lRow, lColumn]:1 + pr:pr)
            else
              Write(lOutputFile, ',');
          end
          else
            Write(lOutputFile, '-,');

        { This is columns of option means }
        if not frmMain.OmitAsinc.IsChecked then
          write(lOutputFile, ',', FOptionMeansArray[lRow, 16]:1 + pr:pr);

        for lColumn := 1 to FMaxItemOptions do
          if FItemStatsArray[lRow, 2] > 0 then
          begin
            if lColumn < lItemScaleArray[lRow, 1] + 1 then
              Write(lOutputFile, ',', FOptionMeansArray[lRow, lColumn]:1 + pr:pr)
            else
              Write(lOutputFile, ',');
          end
          else
            Write(lOutputFile, '-,');

        { This is columns of option SDs }
        if not frmMain.OmitAsinc.IsChecked then
          write(lOutputFile, ',', FOptionSDArray[lRow, 16]:1 + pr:pr);

        for lColumn := 1 to FMaxItemOptions do
          if FItemStatsArray[lRow, 2] > 0 then
          begin
            if lColumn < lItemScaleArray[lRow, 1] + 1 then
              Write(lOutputFile, ',', FOptionSDArray[lRow, lColumn]:1 + pr:pr)
            else
              Write(lOutputFile, ',');
          end
          else
            Write(lOutputFile, '-,');

        SubGroups(lItemInfoArray, lResponseArray, FOmch, lRow, lRankingArray, lReduceArray, lItemScaleArray);

        for lColumn := 1 to FMaxItemOptions do
          If FItemStatsArray[lRow, 2] > 0 then
          begin
            for jfor := 1 to FNumCut do
              if lColumn < lItemScaleArray[lRow, 1] + 1 then
                Write(lOutputFile, ',', FSubGroupPArray[((lColumn - 1) * FNumCut + jfor)]:1 + pr:pr)
              else
                Write(lOutputFile, ',');
          end
          else
            Write(lOutputFile, '-,');

        writeln(lOutputFile); { Move to next lRow of csv file }
      end;

    if frmMain.difcheckbox.IsChecked then
    begin
      writeln(lOutputFile);
      writeln(lOutputFile);
      writeln(lOutputFile, 'Differential Item Functioning Test Results');
      writeln(lOutputFile);
      write(lOutputFile, 'Sequence,Item ID,M-H,M-H SE,z-test,p,Bias Against');
      for lColumn := 1 to StrToIntDef(frmMain.NumDifGroups.Text, 0) do
        write(lOutputFile, ',Score ', IntToStr(lColumn), ' Odds-Ratio');
      writeln(lOutputFile);

      for lRow := 1 to FNumberOfItems do
        if (lItemInfoArray[lRow, 2] = 'Y') and (lItemInfoArray[lRow, 9] = 'D') then
        begin
          Write(lOutputFile, lRow, ',', lItemIDArray[lRow], ',');
          write(lOutputFile, lDifarray[lRow, 1]:1 + pr:pr, ',', lDifarray[lRow, 3]:1 + pr:pr, ',',
            lDifarray[lRow, 4]:1 + pr:pr, ',', lDifarray[lRow, 5]:1 + pr:pr);
          if (lDifarray[lRow, 5] < 0.05) and (lDifarray[lRow, 4] < 0) then
            write(lOutputFile, ',', frmMain.Group2label.Text);
          if (lDifarray[lRow, 5] < 0.05) and (lDifarray[lRow, 4] > 0) then
            write(lOutputFile, ',', frmMain.Group1label.Text);
          if lDifarray[lRow, 5] > 0.05 then
            write(lOutputFile, ',');
          for lColumn := 1 to StrToIntDef(frmMain.NumDifGroups.Text, 0) do
            write(lOutputFile, ',', lDifarray[lRow, 5 + lColumn]:1 + pr:pr);

          writeln(lOutputFile);
        end; // lRow loop
    end; // frmMain.DifCheckBox

    if not frmMain.difcheckbox.IsChecked then
    begin
      writeln(lOutputFile);
      writeln(lOutputFile);
    end;

    writeln(lOutputFile, 'Item statistics - table format');
    writeln(lOutputFile);

    if FIsMult and not FIsRate then
    begin
      if FNumDomains > 1 then
        Write(lOutputFile, 'Sequence,Item ID,Key,Scored,NumOptions,Domain,N,P,Total Rpbis,Total Rbis,Domain Rpbis,',
          'Domain Rbis,Alpha w/o,Flags,Option,Option N,Option Prop,Option Rpbis,Option Rbis,Option Mean,Option SD,')
      else
        Write(lOutputFile, 'Sequence,Item ID,Key,Scored,NumOptions,Domain,N,P,Total Rpbis,Total Rbis,Alpha w/o,Flags,',
          'Option,Option N,Option Prop,Option Rpbis,Option Rbis,Option Mean,Option SD,');
    end; // if mult and no rate

    if FIsMult and FIsRate then
    begin
      if FNumDomains > 1 then
        Write(lOutputFile,
          'Sequence,Item ID,Key,Scored,NumOptions,Domain,N,Item Mean/P,Total Rpbis/R,Total Rbis/Eta,Domain Rpbis/R,',
          'Domain Rbis/Eta,Alpha w/o,Flags,Option,Option N,Option Prop,Option Rpbis,Option Rbis,Option Mean,Option SD,')
      else
        Write(lOutputFile,
          'Sequence,Item ID,Key,Scored,NumOptions,Domain,N,Item Mean/P,Total Rpbis/R,Total Rbis/Eta,Alpha w/o,Flags,',
          'Option,Option N,Option Prop,Option Rpbis,Option Rbis,Option Mean,Option SD,');
    end; // if mult and rate

    if not FIsMult and FIsRate then
    begin
      if FNumDomains > 1 then
        Write(lOutputFile, 'Sequence,Item ID,Key,Scored,NumOptions,Domain,N,Item Mean,Total R,Total Eta,Domain R,',
          'Domain Eta,Alpha w/o,Flags,Option,Option N,Option Prop,Option Rpbis,Option Rbis,Option Mean,Option SD,')
      else
        Write(lOutputFile,
          'Sequence,Item ID,Key,Scored,NumOptions,Domain,N,Item Mean,Total R,Total Eta,Alpha w/o,Flags,',
          'Option,Option N,Option Prop,Option Rpbis,Option Rbis,Option Mean,Option SD,');
    end; // if no mult and rate

    for jfor := 1 to FNumCut do
      write(lOutputFile, 'Grp', IntToStr(jfor), ' P,');

    writeln(lOutputFile);

    { Write FItem stats to output csv file, readable style }
    for lRow := 1 to FNumberOfItems do
      if lItemInfoArray[lRow, 2] <> 'N' then
      begin
        { This is columns pertaining to FItem info }
        if lItemInfoArray[lRow, 2] = 'P' then
          PreText := 'Pretest'
        else
          PreText := 'Yes';
        if lItemInfoArray[lRow, 2] = 'Y' then
          PreStr := formatfloat('0.000', FItemStatsArray[lRow, 5])
        else
          PreStr := 'NA';

        if lNKeyArray[lRow, 1] = 1 then
          Write(lOutputFile, lRow, ',', lItemIDArray[lRow], ',', lItemInfoArray[lRow, 1], ',', PreText, ',',
            lItemScaleArray[lRow, 1], ',', FStoreDomain[lItemScaleArray[lRow, 2]], ',', FItemStatsArray[lRow, 1]:1 + pr
            :pr, ',', FItemStatsArray[lRow, 2]:1 + pr:pr, ',', FItemStatsArray[lRow, 3]:1 + pr:pr, ',',
            FItemStatsArray[lRow, 4]:1 + pr:pr);

        if lNKeyArray[lRow, 1] > 1 then
        begin
          Write(lOutputFile, lRow, ',', lItemIDArray[lRow], ',', lItemInfoArray[lRow, 1]);

          for dfor := 1 to lNKeyArray[lRow, 1] - 1 do
            write(lOutputFile, ' ', UpperCase(lItemInfoArray[lRow, 3 + dfor]));

          write(lOutputFile, ',', PreText, ',', lItemScaleArray[lRow, 1], ',', FStoreDomain[lItemScaleArray[lRow, 2]],
            ',', FItemStatsArray[lRow, 1]:1 + pr:pr, ',', FItemStatsArray[lRow, 2]:1 + pr:pr, ',',
            FItemStatsArray[lRow, 3]:1 + pr:pr, ',', FItemStatsArray[lRow, 4]:1 + pr:pr);
        end;

        if FNumDomains > 1 then
          write(lOutputFile, ',', FItemStatsArray[lRow, 7]:1 + pr:pr, ',', FItemStatsArray[lRow, 8]:1 + pr:pr);

        write(lOutputFile, ',', PreStr, ',');

        lCheckKey := False;

        { Flags }
        for jfor := 1 to lItemScaleArray[lRow, 1] do
          if (lItemInfoArray[lRow, 3] = 'M') and (FItemStatsArray[lRow, 6] <> 0) then
            if FItemStatsArray[lRow, 3] < FOptionRpbisArray[lRow, jfor] then
              lCheckKey := True;

        if lCheckKey then
          write(lOutputFile, FFlagArray[1], ' ');

        if lItemInfoArray[lRow, 3] = 'M' then
        begin
          if FItemStatsArray[lRow, 2] < FMinP then
            write(lOutputFile, FFlagArray[2], ' ');
          if FItemStatsArray[lRow, 2] > FMaxP then
            write(lOutputFile, FFlagArray[3], ' ');
        end;

        if lItemInfoArray[lRow, 3] <> 'M' then
        begin
          if FItemStatsArray[lRow, 2] < FMinMean then
            write(lOutputFile, FFlagArray[6], ' ');
          if FItemStatsArray[lRow, 2] > FMaxMean then
            write(lOutputFile, FFlagArray[7], ' ');
        end;

        if FItemStatsArray[lRow, 3] < FMinR then
          write(lOutputFile, FFlagArray[4], ' ');
        if FItemStatsArray[lRow, 3] > FMaxR then
          write(lOutputFile, FFlagArray[5], ' ');
        if (lItemInfoArray[lRow, 9] = 'D') and (frmMain.difcheckbox.IsChecked) then
          if lDifarray[lRow, 5] < 0.05 then
            write(lOutputFile, FFlagArray[8], ' ');

        SubGroups(lItemInfoArray, lResponseArray, FOmch, lRow, lRankingArray, lReduceArray, lItemScaleArray);
        // added 4/12/12

        { output first lRow of option stats }
        if (FRespAsInt = True) and (lItemInfoArray[lRow, 3] <> 'P') then
          write(lOutputFile, ',1,');
        if lItemInfoArray[lRow, 3] = 'P' then
          write(lOutputFile, ',0,');
        if (FRespAsInt = False) and (lItemInfoArray[lRow, 3] <> 'P') then
          write(lOutputFile, ',', chr(65), ',');
        Write(lOutputFile, FOptionCountsArray[lRow, 1]);
        // Write(lOutputFile, ',,1,', FOptionCountsArray[lRow,1]);
        Write(lOutputFile, ',', FOptionPropsArray[lRow, 1]:1 + pr:pr);

        if FItemStatsArray[lRow, 2] > 0 then
        begin
          Write(lOutputFile, ',', FOptionRpbisArray[lRow, 1]:1 + pr:pr);
          Write(lOutputFile, ',', FOptionRbisArray[lRow, 1]:1 + pr:pr);
          Write(lOutputFile, ',', FOptionMeansArray[lRow, 1]:1 + pr:pr);
          write(lOutputFile, ',', FOptionSDArray[lRow, 1]:1 + pr:pr);

          for jfor := 1 to FNumCut do
            Write(lOutputFile, ',', FSubGroupPArray[jfor]:1 + pr:pr);
        end
        else
          Write(lOutputFile, ',-,-,-,-,-,-');

        if lItemInfoArray[lRow, 3] = 'P' then
          SCO := 1
        else
          SCO := 0;

        if ConvertResponseToInteger(lItemInfoArray[lRow, 1], lItemInfoArray, lRow, lItemScaleArray[lRow, 1]) = 1 - SCO
        then
          Write(lOutputFile, ',**KEY**');

        writeln(lOutputFile);

        { output rest of option stats }
        for lColumn := 2 to lItemScaleArray[lRow, 1] + 2 do
        begin
          Write(lOutputFile, ',,,,,,,,,,');

          if FNumDomains > 1 then
            write(lOutputFile, ',,');
          if lColumn <= lItemScaleArray[lRow, 1] then
          begin
            if (FRespAsInt = True) and (lItemInfoArray[lRow, 3] <> 'P') then
              write(lOutputFile, ',,', lColumn);
            if lItemInfoArray[lRow, 3] = 'P' then
              write(lOutputFile, ',,', lColumn - 1);
            if (FRespAsInt = False) and (lItemInfoArray[lRow, 3] <> 'P') then
              write(lOutputFile, ',,', chr(lColumn + 64));
          end;

          if lColumn = lItemScaleArray[lRow, 1] + 1 then
          begin
            write(lOutputFile, ',,Omit', ',', FOptionCountsArray[lRow, 16]);

            if not frmMain.OmitAsinc.IsChecked then
              write(lOutputFile, ',', FOptionPropsArray[lRow, 16]:1 + pr:pr)
            else
              write(lOutputFile, ',-');

            if FItemStatsArray[lRow, 2] > 0 then
            begin
              if not frmMain.OmitAsinc.IsChecked then
                Write(lOutputFile, ',', FOptionRpbisArray[lRow, 16]:1 + pr:pr)
              else
                write(lOutputFile, ',-');

              if not frmMain.OmitAsinc.IsChecked then
                Write(lOutputFile, ',', FOptionRbisArray[lRow, 16]:1 + pr:pr)
              else
                write(lOutputFile, ',-');

              Write(lOutputFile, ',', FOptionMeansArray[lRow, 16]:1 + pr:pr);
              Write(lOutputFile, ',', FOptionSDArray[lRow, 16]:1 + pr:pr);
            end
            else
              write(lOutputFile, ',-,-,-,-,-,-');

            writeln(lOutputFile);
          end;

          if lColumn = lItemScaleArray[lRow, 1] + 2 then
          begin
            write(lOutputFile, ',,Not Reached', ',', FOptionCountsArray[lRow, 17], ',-');

            if FItemStatsArray[lRow, 2] > 0 then
            begin
              write(lOutputFile, ',-,-');
              Write(lOutputFile, ',', FOptionMeansArray[lRow, 17]:1 + pr:pr);
              Write(lOutputFile, ',', FOptionSDArray[lRow, 17]:1 + pr:pr);
              write(lOutputFile, ',-,-,-');
            end
            else
              write(lOutputFile, ',-,-,-,-,-,-');

            writeln(lOutputFile);
          end;

          if lColumn <= lItemScaleArray[lRow, 1] then
          begin
            Write(lOutputFile, ',', FOptionCountsArray[lRow, lColumn]);
            Write(lOutputFile, ',', FOptionPropsArray[lRow, lColumn]:1 + pr:pr);

            If FItemStatsArray[lRow, 2] > 0 then
            begin
              Write(lOutputFile, ',', FOptionRpbisArray[lRow, lColumn]:1 + pr:pr);
              Write(lOutputFile, ',', FOptionRbisArray[lRow, lColumn]:1 + pr:pr);
              Write(lOutputFile, ',', FOptionMeansArray[lRow, lColumn]:1 + pr:pr);
              Write(lOutputFile, ',', FOptionSDArray[lRow, lColumn]:1 + pr:pr);
              for jfor := 1 to FNumCut do
                Write(lOutputFile, ',', FSubGroupPArray[jfor + (lColumn - 1) * FNumCut]:1 + pr:pr);
            end
            else
              Write(lOutputFile, ',-,-,-,-,-,-');

            if ConvertResponseToInteger(lItemInfoArray[lRow, 1], lItemInfoArray, lRow, lItemScaleArray[lRow, 1])
              = lColumn - SCO then
              Write(lOutputFile, ',**KEY**');

            jfor := 1;
            if lNKeyArray[lRow, 1] > 1 then
              if ConvertResponseToInteger(lItemInfoArray[lRow, 3 + jfor], lItemInfoArray, lRow, lItemScaleArray[lRow, 1]
                ) = lColumn - SCO then
              begin
                Write(lOutputFile, ',**KEY' + IntToStr(succ(jfor)) + '**');
                // inc(jfor); // Never used
              end;

            writeln(lOutputFile);
          end;
        end;

        writeln(lOutputFile);
      end;

    { -------------------------Raw frequency table--------------------------- }
    writeln(lOutputFile, 'Frequency Table for Total Score for all the Scored Items');
    writeln(lOutputFile);
    writeln(lOutputFile, 'Raw Score,Freq,Prop,Cum Freq,Cum Prop');

    for lRow := 1 to FNumUniqueFreq do
    begin
      write(lOutputFile, round(lRawFrequency[lRow - 1, 1]), ',', round(lRawFrequency[lRow - 1, 2]));
      write(lOutputFile, ',', (lRawFrequency[lRow - 1, 2] / FNumberOfExaminees):1 + 7:7);
      write(lOutputFile, ',', round(lRawFrequency[lRow - 1, 3]));
      write(lOutputFile, ',', (lRawFrequency[lRow - 1, 3] / FNumberOfExaminees):1 + 7:7);
      writeln(lOutputFile);
    end;
  end;

  function BBOMatrixOutputFile(): Boolean;
  var
    lBBO_File: TextFile;
    lEIC_File: TextFile;
    lEEIC_File: TextFile;
    lColumn: Integer;
    lRow: Integer;
  begin
    Result := True;

    if lInputOk then
    begin
      { ----Prepare BBO File ---- }
      lOutputFileName := ChangeFileExt(frmMain.eOutputFile.Text, '') + ' BBO-matrix';
      lOutputFileName := ChangeFileExt(lOutputFileName, '.csv');
      lOutputFileName := OutputDir + lOutputFileName;
      AssignFile(lBBO_File, lOutputFileName);

      try
        ReWrite(lBBO_File);
      except
        lInputOk := False;
        ErrorMessage := ('The BBO-matrix CSV file is already open in another program and cannot be saved by Iteman.' +
          #13 + 'Please close the file, then hit Retry to continue the analysis.');
        Result := False;
        Exit;
      end;

      { ----Prepare EIC File ---- }
      lOutputFileName := ChangeFileExt(frmMain.eOutputFile.Text, '') + ' EIC-matrix';
      lOutputFileName := ChangeFileExt(lOutputFileName, '.csv');
      lOutputFileName := OutputDir + lOutputFileName;
      AssignFile(lEIC_File, lOutputFileName);

      try
        ReWrite(lEIC_File);
      except
        lInputOk := False;
        ErrorMessage := ('The EIC-matrix CSV file is already open in another program and cannot be saved by Iteman.' +
          #13 + 'Please close the file, then hit Retry to continue the analysis.');
        Result := False;
        Exit;
      end;

      { ----Prepare EEIC File ---- }
      lOutputFileName := ChangeFileExt(frmMain.eOutputFile.Text, '') + ' EEIC-matrix';
      lOutputFileName := ChangeFileExt(lOutputFileName, '.csv');
      lOutputFileName := OutputDir + lOutputFileName;
      AssignFile(lEEIC_File, lOutputFileName);

      try
        ReWrite(lEEIC_File);
      except
        lInputOk := False;
        ErrorMessage := ('The EEIC-matrix CSV file is already open in another program and cannot be saved by Iteman.' +
          #13 + 'Please close the file, then hit Retry to continue the analysis.');
        Result := False;
        Exit;
      end; // lInputOk

      { -- fill in BB matrix -- }
      write(lBBO_File, ', Flag');

      // Write lColumn headers
      for lColumn := 1 to FNumberOfExaminees do
        write(lBBO_File, ',', lExamineeIDArray[lColumn]);

      writeln(lBBO_File);

      for lColumn := 1 to FNumberOfExaminees do
        write(lEIC_File, ',', lExamineeIDArray[lColumn]);

      writeln(lEIC_File);

      for lColumn := 1 to FNumberOfExaminees do
        write(lEEIC_File, ',', lExamineeIDArray[lColumn]);

      writeln(lEEIC_File);

      // Fill in matrix files
      for lRow := 1 to FNumberOfExaminees do
      begin
        write(lBBO_File, lExamineeIDArray[lRow], ',', FBBOFlags[lRow]);

        for lColumn := 1 to FNumberOfExaminees do
        begin
          if lColumn < lRow then
            write(lBBO_File, ',', FBBOindexArray[lRow, lColumn]:3:4)
          else
            write(lBBO_File, ',');

          if lColumn < lRow then
            write(lEIC_File, ',', FEICArray[lRow, lColumn])
          else
            write(lEIC_File, ',');

          if lColumn < lRow then
            write(lEEIC_File, ',', FEEICArray[lRow, lColumn])
          else
            write(lEEIC_File, ',');
        end; // for lColumn loop

        // go to next line
        writeln(lBBO_File);
        writeln(lEIC_File);
        writeln(lEEIC_File);
      end; // lRow

      CloseFile(lBBO_File);
      CloseFile(lEIC_File);
      CloseFile(lEEIC_File);
    end; // BB output
  end;

  procedure DOCXResultFileCreating();
  var
    brange, lRow, k: Integer;
    FItem: Integer;

    { -------------compute number of bins for hist if > 20 items in domain-------------- }
    function calcbin(Max: Integer): Integer;
    var
      ij, temp1: Integer;
    begin
      nbin := 0;
      for ij := 9 to 20 do
      begin
        temp1 := ((brange) mod ij);
        if temp1 = 0 then
          nbin := ij + 1;
      end;
      if nbin = 0 then
        nbin := 20 + 1; // if 0 then no even divisor...
    end;

  begin
    // _____________ Establish Export Headers ________________

    frmMain.ProgressLabel.Text := 'Producing DOCX Output';

    if FNumFlagItems = 0 then
      tableno := 4
    else
      tableno := 5;

    if InitExport then
      try
        AddHeaders(lRunTitle, FTitleBitmap, FIntroductionBitmap, frmMain.Is_Demo);
        AddSpecifications(FSpecBitmap, frmMain.GetInputFiles(data, ctrl), frmMain.GetOutputFiles(data, ctrl),
          GetTableSpec, UseMRF);
        PrintFirstSummary(prec, lItemInfoArray, lItemIDArray, lItemScaleArray, lDifarray);

        { Write description of Figure 1 to docx outpout }
        AddText('');
        AddText(PWideChar
          ('Figure 1 displays the distribution of the raw scores for the scored items across all domains. ' + 'Table ' +
          IntToStr(tableno) + ' displays the frequency distribution for total score shown in Figure 1.'));

        { Insert graphic on first page }
        if FTestStatsArray[1, 5] <= 19 then
          nbin := round(FTestStatsArray[1, 5] - FTestStatsArray[1, 4] + 1);

        brange := round(FTestStatsArray[1, 5] - FTestStatsArray[1, 4]);
        if FTestStatsArray[1, 5] >= 20 then
          calcbin(round(FTestStatsArray[1, 5]));
        CalcGF(FTestStatsArray[1, 4], FTestStatsArray[1, 5], 1, lScoreArray);
        DrawDistGraph(FTestStatsArray[1, 4], FTestStatsArray[1, 5], round(FTestStatsArray[1, 1]), FNumberOfExaminees, 1,
          FGraphBitmap, 'Total Score');
        AddImage('Figure 1: Total score for the scored items', FGraphBitmap);

        // FExport.AddText('');
        AddTable('Table ' + IntToStr(tableno) + ': Frequency Distribution for Total Score', nil);
        AddText('');
        tableno := tableno + 1;

        { TODO : refactoring! }
        if nbin <= 10 then
          PrintFreq10('0', FTestStatsArray[1, 4], FTestStatsArray[1, 5]);
        if (nbin > 10) and (nbin < 16) then
          PrintFreq15('0', FTestStatsArray[1, 4], FTestStatsArray[1, 5]);
        if nbin > 15 then
          PrintFreq20('0', FTestStatsArray[1, 4], FTestStatsArray[1, 5]);
        AddText('');

        if FNumDomains > 1 then
          for lRow := 1 to FNumDomains do
          begin
            inc(figureno);

            AddText('');
            AddText(PWideChar('Figure ' + IntToStr(figureno) + ' displays the distribution of the raw scores for ' +
              FStoreDomain[lRow] + '. ' + 'Table ' + IntToStr(tableno) +
              ' displays the frequency distribution of domain scores shown in Figure ' + IntToStr(figureno) + '.'));
            { Write description of Figure 1 to docx outpout }
            AddText('');
            AddText('');

            { Insert graphic on first page }
            if FDomainStatsArray[lRow, 5] <= 19 then
              nbin := round(FDomainStatsArray[lRow, 5] - FDomainStatsArray[lRow, 4] + 1);

            brange := round(FDomainStatsArray[lRow, 5] - FDomainStatsArray[lRow, 4]);

            if FDomainStatsArray[lRow, 5] >= 20 then
              calcbin(round(FDomainStatsArray[lRow, 5]));

            CalcGF(FDomainStatsArray[lRow, 4], FDomainStatsArray[lRow, 5], lRow + 1, lScoreArray);
            DrawDistGraph(FDomainStatsArray[lRow, 4], FDomainStatsArray[lRow, 5], round(FDomainStatsArray[lRow, 1]),
              FNumberOfExaminees, 1, FGraphBitmap, FStoreDomain[lRow]);
            AddImage('Figure ' + IntToStr(lRow + 1) + ': Raw scores for ' + FStoreDomain[lRow], FGraphBitmap);

            AddText('');
            AddTable('Table ' + IntToStr(tableno) + ': Frequency Distribution for ' + FStoreDomain[lRow], nil);
            inc(tableno);

            if nbin <= 10 then
              PrintFreq10('0', FDomainStatsArray[lRow, 4], FDomainStatsArray[lRow, 5]);
            if (nbin > 10) and (nbin < 16) then
              PrintFreq15('0', FDomainStatsArray[lRow, 4], FDomainStatsArray[lRow, 5]);
            if nbin > 15 then
              PrintFreq20('0', FDomainStatsArray[lRow, 4], FDomainStatsArray[lRow, 5]);
          end;

        { ---------------------- Subscore Correlation Feb.2013 ---------------------- }
        if FNumDomains > 1 then
        begin
          AddText('');
          AddText(PWideChar('Table ' + IntToStr(tableno) + ' displays the correlations of domain scores. '));

          SetLength(tbMatrix, FNumDomains + 1, FNumDomains + 1);
          tbMatrix[0, 0] := 'Domain';
          for k := 1 to FNumDomains do
            tbMatrix[0, k] := FStoreDomain[k];
          for lRow := 1 to FNumDomains do
            tbMatrix[lRow, 0] := FStoreDomain[lRow];

          for lRow := 1 to FNumDomains do
            for k := 1 to FNumDomains do
              tbMatrix[lRow, k] := formatfloat(prec, FDomainCorrArray[lRow, k]);

          AddTable('Table ' + IntToStr(tableno) + ': Correlations for Domain Scores', tbMatrix);
          tableno := tableno + 1;
        end;

        if lScoredMult then
        begin
          figureno := figureno + 1;
          // FExport.AddText('');
          AddText(PWideChar('Figure ' + IntToStr(figureno) +
            ' displays the distribution of the P values for the dichotomously scored items (correct/incorrect). ' +
            'Table ' + IntToStr(tableno) + ' displays the frequency distribution of the P values shown in Figure ' +
            IntToStr(figureno) + '.'));
          AddText('');
          nbin := 10;
          brange := 1;
          CalcItemGF(0, 1, 1, FNumMult[1], 'M', lItemInfoArray);

          DrawDistGraph(0, 1, FNumMult[1], FNumMult[1], 2, FGraphBitmap, 'P value');
          AddImage('Figure ' + IntToStr(figureno) + ': P values for the scored items', FGraphBitmap);

          AddText('');
          AddTable('Table ' + IntToStr(tableno) + ': Frequency Distribution for the P values', nil);
          tableno := tableno + 1;
          if nbin = 10 then
            PrintFreq10('0.0', -1, -1);
          if (nbin > 10) and (nbin < 16) then
            PrintFreq15('0.0', -1, -1);
          if nbin > 15 then
            PrintFreq20('0.0', -1, -1);
        end;

        if scoredrate then
        begin
          figureno := figureno + 1;
          AddText('');
          AddText(PWideChar('Figure ' + IntToStr(figureno) + ' displays the distribution of the Item Means ' +
            '(rating scale or partial credit items only).  Table ' + IntToStr(tableno) +
            ' displays the frequency distribution of the Item Means shown in Figure ' + IntToStr(figureno) + '.'));
          AddText('');
          nbin := 10;
          if IsZero = True then
            i := 0
          else
            i := 1;
          if IsZero = True then
            j := 1
          else
            j := 2;
          if maxcategory > 2 then
            FMinItem1 := floor(FMinItem1);
          if maxcategory > 2 then
            FMaxItem1 := ceil(FMaxItem1);
          if (FMaxItem1 - FMinItem1) < 1 then
            brange := 1
          else
            brange := round(FMaxItem1 - FMinItem1);
          if maxcategory = 2 then
            FMinItem1 := i;
          if maxcategory = 2 then
            FMaxItem1 := j;
          CalcItemGF(FMinItem1, FMaxItem1, 1, FNumRate[1], 'R', lItemInfoArray);

          DrawDistGraph(FMinItem1, FMaxItem1, FNumRate[1], FNumRate[1], 3, FGraphBitmap, 'Item mean');
          AddImage('Figure ' + IntToStr(figureno) + ': Item Means for the scored items', FGraphBitmap);
          AddText('');

          AddTable('Table ' + IntToStr(tableno) + ': Frequency Distribution for the Item Means', nil);
          tableno := tableno + 1;
          if nbin = 10 then
            PrintFreq10('0.0', -1, -1);
          if (nbin > 10) and (nbin < 16) then
            PrintFreq15('0.0', -1, -1);
          if nbin > 15 then
            PrintFreq20('0.0', -1, -1);
        end;

        if lScoredMult then
        begin
          Rrange := 1;
          figureno := figureno + 1;
          AddText('');
          AddText(PWideChar('Figure ' + IntToStr(figureno) +
            ' displays the distribution of the Point-Biserial Correlations ' +
            'for the dichotomously scored items (correct/incorrect).  Table ' + IntToStr(tableno) +
            ' displays the frequency distribution of the Point-Biserial correlations shown in Figure ' +
            IntToStr(figureno) + '.'));
          AddText('');
          if (FMinRPBis < 0) and (FMinRPBis > -0.1) then
            Rrange := 1.1;
          if (FMinRPBis < -0.1) and (FMinRPBis > -0.2) then
            Rrange := 1.2;
          if (FMinRPBis < -0.2) and (FMinRPBis > -0.3) then
            Rrange := 1.3;
          if (FMinRPBis < -0.3) and (FMinRPBis > -0.4) then
            Rrange := 1.4;
          if (FMinRPBis < -0.4) and (FMinRPBis > -0.5) then
            Rrange := 1.5;
          if (FMinRPBis < -0.5) and (FMinRPBis > -0.6) then
            Rrange := 1.6;
          if (FMinRPBis < -0.6) and (FMinRPBis > -0.7) then
            Rrange := 1.7;
          if (FMinRPBis < -0.7) and (FMinRPBis > -0.8) then
            Rrange := 1.8;
          if (FMinRPBis < -0.8) and (FMinRPBis > -0.9) then
            Rrange := 1.9;
          if (FMinRPBis < -0.9) and (FMinRPBis > -1.0) then
            Rrange := 2.0;
          nbin := round(Rrange * 10);
          brange := 1;
          CalcItemGF(1 - Rrange, 1, 4, FNumMult[1], 'M', lItemInfoArray);

          DrawDistGraph(1 - Rrange, 1, FNumMult[1], FNumMult[1], 4, FGraphBitmap, 'Rpbis');
          AddImage('Figure ' + IntToStr(figureno) + ': Rpbis for the scored items', FGraphBitmap);

          AddTable('Table ' + IntToStr(tableno) + ': Frequency Distribution for the Rpbis', nil);
          tableno := tableno + 1;
          if nbin = 10 then
            PrintFreq10('0.0', -1, -1);
          if (nbin > 10) and (nbin < 16) then
            PrintFreq15('0.0', -1, -1);
          if nbin > 15 then
            PrintFreq20('0.0', -1, -1);
        end;

        if scoredrate then
        begin
          Rrange := 1;
          figureno := figureno + 1;
          AddText('');
          AddText(PWideChar('Figure ' + IntToStr(figureno) +
            ' displays the distribution of the Item-Total Pearson r Correlations ' +
            'for polytomously scored items (rating scale or partial credit). Table ' + IntToStr(tableno) +
            ' displays the frequency distribution of Pearsons r correlations shown in Figure ' +
            IntToStr(figureno) + '.'));
          AddText('');
          if (FMinCorr < 0) and (FMinCorr > -0.1) then
            Rrange := 1.1;
          if (FMinCorr < -0.1) and (FMinCorr > -0.2) then
            Rrange := 1.2;
          if (FMinCorr < -0.2) and (FMinCorr > -0.3) then
            Rrange := 1.3;
          if (FMinCorr < -0.3) and (FMinCorr > -0.4) then
            Rrange := 1.4;
          if (FMinCorr < -0.4) and (FMinCorr > -0.5) then
            Rrange := 1.5;
          if (FMinCorr < -0.5) and (FMinCorr > -0.6) then
            Rrange := 1.6;
          if (FMinCorr < -0.6) and (FMinCorr > -0.7) then
            Rrange := 1.7;
          if (FMinCorr < -0.7) and (FMinCorr > -0.8) then
            Rrange := 1.8;
          if (FMinCorr < -0.8) and (FMinCorr > -0.9) then
            Rrange := 1.9;
          if (FMinCorr < -0.9) and (FMinCorr > -1.0) then
            Rrange := 2.0;
          nbin := round(Rrange * 10);
          brange := 1;
          CalcItemGF(1 - Rrange, 1, 4, FNumRate[1], 'R', lItemInfoArray);

          DrawDistGraph(1 - Rrange, 1, FNumRate[1], FNumRate[1], 4, FGraphBitmap, 'R');
          AddImage('Figure ' + IntToStr(figureno) + ': R for the scored items', FGraphBitmap);

          AddText('');

          AddTable('Table ' + IntToStr(tableno) + ': Frequency Distribution for Pearsons r', nil);
          tableno := tableno + 1;
          if nbin = 10 then
            PrintFreq10('0.0', -1, -1);
          if (nbin > 10) and (nbin < 16) then
            PrintFreq15('0.0', -1, -1);
          if nbin > 15 then
            PrintFreq20('0.0', -1, -1);
        end;

        if lScoredMult then
        begin
          nbin := 10;
          figureno := figureno + 1;
          AddText('');
          AddText(PWideChar('Figure ' + IntToStr(figureno) + ' displays the scatterplot of P (difficulty) by Rpbis ' +
            '(discrimination) for the dichotomously scored items (correct/incorrect).'));
          brange := 1;
          DrawScatterplot(0, 1, FNumMult[1], FNumberOfItems, 1, FGraphBitmap, lItemInfoArray, 'M', 'P value');
          AddImage('Figure ' + IntToStr(figureno) + ': P by Rpbis', FGraphBitmap);

          if figureno mod 2 = 0 then
            AddText('');
        end;

        if scoredrate then
        begin
          nbin := 10;
          figureno := figureno + 1;
          AddText('');
          if lScoredMult = True then
            AddText('');
          AddText(PWideChar('Figure ' + IntToStr(figureno) +
            ' displays the scatterplot of the Item Means (difficulty) by R ' +
            '(discrimination) for polytomously scored items (rating scale or partial credit).'));
          if IsZero = True then
            brange := 0
          else
            brange := 1;
          if lAllPC = True then
            DrawScatterplot(brange, maxcategory - 1, FNumRate[1], FNumberOfItems, 2, FGraphBitmap, lItemInfoArray, 'R',
              'Item Mean')
          else
            DrawScatterplot(brange, maxcategory, FNumRate[1], FNumberOfItems, 2, FGraphBitmap, lItemInfoArray, 'R',
              'Item Mean');
          // lAllPC is all partial credit (begins at 0)
          ;
          AddImage('Figure ' + IntToStr(figureno) + ': Item Means by R', FGraphBitmap);
        end;

        if not scoredrate then
        begin
          AddText('');
          tmp := 'Figure ' + IntToStr(figureno + 1) +
            ' displays a graph of the Conditional Standard Error of Measurement (CSEM) Formula IV. ';
          if (rawcut >= 0) and (rawcut <= FTestStatsArray[1, 1]) then
            if frmMain.classcheckbox.IsChecked then
              tmp := tmp + ' The CSEM at the cutscore of ' + formatfloat(prec, rawcut) + ' equaled ' +
                formatfloat(prec, sqrt(lCsemArray[round(rawcut), 2])) + '.';
          if (rawcut < 0) or (rawcut > FTestStatsArray[1, 1]) then
            tmp := tmp + ' The scaled number correct cut-score of ' + formatfloat(prec, rawcut) +
              ' is outside of the number correct score range, the CSEM cannot be computed.';
          if rawcut < 0 then
            tmp := tmp + ' All examinees were classified into the group: ' + frmMain.LowLabel.Text + '.';
          if rawcut > FTestStatsArray[1, 1] then
            tmp := tmp + ' All examinees were classified into the group: ' + frmMain.HighLabel.Text + '.';
          AddText(PWideChar(tmp));
          DrawCsemGraph(round(FTestStatsArray[1, 1]), FGraphBitmap, 'CSEM', lCsemArray);
          AddImage('Figure ' + IntToStr(figureno + 1) + ': CSEM', FGraphBitmap);
        end;

        { docx description of FItem-by-FItem results }
        AddText('');
        // is this an extra blank line above the quantile plot?
        AddItemByItemDesc(frmMain.plotsbox.IsChecked, frmMain.tablebox.IsChecked);

        { Write each FItem's statistics to docx output no longer AS A RICHVIEW TABLE }
        for FItem := 1 to FNumberOfItems do
          if lItemInfoArray[FItem, 2] <> 'N' then
          begin
            ticker := ticker + 1;
            AddText('<#Page#>'); // page break
            SubGroups(lItemInfoArray, lResponseArray, FOmch, FItem, lRankingArray, lReduceArray, lItemScaleArray);
            { Draw FItem graph and paste in }
            If frmMain.plotsbox.IsChecked then
            begin
              DrawItemGraph(FGraphBitmap, FItem, lItemScaleArray, lItemIDArray, lItemInfoArray[FItem, 1],
                lItemInfoArray[FItem, 3]);
              AddImage('', FGraphBitmap);
            end;

            if not frmMain.plotsbox.IsChecked then
              AddText(''); // so page breaks go in correct spot

            { Insert page break between all items }
            // if plotsbox.IsChecked then FExport.AddText('<#Page#>');     //pagebreak
            PrintItemTables(prec, FItem, lItemInfoArray, lItemIDArray, lItemScaleArray, lNKeyArray,
              lItemInfoArray[FItem, 3], lDifarray);

            frmMain.ProgressBar.Value := frmMain.ProgressBar.Value + 1;
            Application.ProcessMessages;

            { Save all docx file }
            if (ticker mod itemcap = 0) or ((FItem = FNumberOfItems) and (FItem >= itemcap)) then
            begin
              if ticker mod itemcap = 0 then
                start1 := FItem;
              if ticker mod itemcap = 0 then
                itemcap := 250
              else
                itemcap := 0;
              if nreport = 0 then
                fino := ' Items 1' + '-' + IntToStr(FItem);
              if nreport <> 0 then
                fino := ' Items ' + IntToStr(start1 + 1 - itemcap) + '-' + IntToStr(FItem);
              OutputDir := path;
              if UseMRF then
                frmMain.docxoutput := ChangeFileExt(path + Filename + fino, '.docx')
              else
              begin
                frmMain.docxoutput := ChangeFileExt(outp, '');
                frmMain.docxoutput := ChangeFileExt(frmMain.docxoutput + fino, '.docx');
              end;

              try
                SaveToFile(PWideChar(frmMain.docxoutput));
              except
                on E: Exception do
                  // ErrorMessage:=(E.Message, mtError, [mbOK], 0);
                  ErrorMessage := E.Message;
              end;

              ReInit;
              nreport := nreport + 1;
            end; // FItem mod itemcap statement
          end; // items loop
      finally
        SaveToFile(PWideChar(frmMain.docxoutput));
        DoneExport;
      end;
    { ________________________________________________________ }
    if mrf1 <> 'GUI' then
    begin
      frmMain.mLog.lines.Add('The analysis was completed successfully.');
      frmMain.mLog.lines.Add(IntToStr(FNumberOfItems) + ' items and ' + IntToStr(FNumberOfExaminees) +
        ' examinees were processed');
      frmMain.mLog.lines.Add('');
      Application.ProcessMessages;
    end;
  end;

var
  lColumn: Integer;
  lRow: Integer;
begin
  lTabChar := chr(0);
  lIDBegin := 0;
  lIsDelim := False;
  delim1 := chr(0);
  lGroupCodeColumn := 0;

  try
    FIsMult := False;

    lMismatch := False;
    lAllPilot := True;
    // if all items are pilot, then program will stop and give an error
    lAllPC := True;
    lZeroOk := True;
    FAllOk := True; { TODO : Consider deleting. Should not be mandatory }
    lInputOk := True;
    lColumnOk := True;

    FIsRate := False;
    scoredrate := False;
    lScoredMult := False;
    lKeyOk := True; // //error trapping missing values in FKey added Feb.2013

    itemcap := 250;
    ticker := 0;
    nreport := 0;
    UseMRF := not(mrf1 = 'GUI');
    pr := StrToIntDef(frmMain.Precision.Text, 0);

    for i := 1 to 4 do
      FNumMult[i] := 0;
    for i := 1 to 4 do
      FNumRate[i] := 0;
    for i := 1 to 100 do
      FDomMult[i] := 0;
    for i := 1 to 100 do
      FDomRate[i] := 0;
    for i := 1 to Length(datacontrolerror) do
      datacontrolerror[i] := False;
    for i := 1 to Length(patharray) do
      patharray[i] := #0;
    for i := 1 to Length(readformat) do
      readformat[i] := #0;

    maxcategory := 0;
    figureno := 1;

    // check to see if all input has been taken care of
    Result := CheckSettings();
    if not Result then
      Exit;

    { Assign files moved to trap errors in input file accessibility }
    Result := CheckDataMatrixAndControlFiles();
    if not Result then
      Exit;

    { ---- Check control and data have the same number of items  ------- }
    Result := CheckControlAndDataItemsCount();
    if not Result then
      Exit;

    Result := CheckOutputAndScoreFile();
    if not Result then
      Exit;

    if lInputOk then
    begin
      Reset(lDataMatrixFile); // resetting

      { Read input values }
      lRunTitle := frmMain.RunTitleBox.Text;
      FNumFlagItems := 0;
      lOmitString := UpperCase(frmMain.omitcharbox.Text);
      lNAString := UpperCase(frmMain.nacharbox.Text);
      FOmitChar := lOmitString[1];
      FNAChar := lNAString[1];
      FMinP := GetFloatValue(frmMain.MinPBox.Text);
      // StrToFloatDef(frmMain.MinPBox.Text, 0);
      FMaxP := GetFloatValue(frmMain.MaxPBox.Text);
      // StrToFloatDef(frmMain.MaxPBox.Text, 0);
      FMinR := GetFloatValue(frmMain.MinRpbisBox.Text);
      // StrToFloatDef(frmMain.MinRpbisBox.Text, 0);
      FMaxR := GetFloatValue(frmMain.MaxRpbisBox.Text);
      // StrToFloatDef(frmMain.MaxRpbisBox.Text, 0);
      FMaxMean := GetFloatValue(frmMain.MaxMeanBox.Text);
      // StrToFloatDef(frmMain.MaxMeanBox.Text, 0);
      FMinMean := GetFloatValue(frmMain.MinMeanBox.Text);
      // StrToFloatDef(frmMain.MinMeanBox.Text, 0);
      lTypeOk := True;
      lAltOk := True;
      lDomOk := True;
      lEndLine := True;

      { Slightly change bounds to allow values equal to original bounds }
      FMinP := FMinP - 0.00001;
      FMaxP := FMaxP + 0.00001;
      FMinR := FMinR - 0.00001;
      FMaxR := FMaxR + 0.00001;
      FMinMean := FMinMean - 0.00001;
      FMaxMean := FMaxMean + 0.00001;
      FRespOK := True; // error trapping unidentified response characters

      { if nex > 1 then
        previd:= lExamineeIDColumns; }
      // ��� ������� �� ����� ������, �.�. nex ����� �� �������.

      { --- Find the number of rows (examinees) in input --- }
      FNumberOfExaminees := 0;

      while not EOF(lDataMatrixFile) do
      begin
        ReadLn(lDataMatrixFile, lChar);

        if lChar <> chr(13) then
          inc(FNumberOfExaminees);
      end;

      Reset(lDataMatrixFile);
      if frmMain.cbDataMatrixFile.IsChecked then
        FNumberOfExaminees := FNumberOfExaminees - 4;

      { --- Demo Limitation --- }
      if frmMain.Is_Demo and (FNumberOfExaminees > 100) then
        FNumberOfExaminees := 100;

      if not frmMain.cbDataMatrixFile.IsChecked then
      begin { --- Find the number of rows (items) in FItem file --- }
        FNumberOfItems := 0;

        while not EOF(lItemInfoFile) do
        begin
          ReadLn(lItemInfoFile, lChar);
          // chr(13) is return, so blank lines are not treated as items

          if (lChar <> chr(13)) then
            inc(FNumberOfItems);
        end;

        Reset(lItemInfoFile);
      end;

      { --- Demo Limitation --- }
      if frmMain.Is_Demo and (FNumberOfItems > 100) then
        FNumberOfItems := 100;

      FNumCut := StrToIntDef(frmMain.CutPoint.Text, 0);

      // Read in data and set up arrays
      if not frmMain.cbDataMatrixFile.IsChecked then
      begin
        // Define all dynamic arrays - with some extra room just in case
        SetLength(lResponseArray, (FNumberOfExaminees + 10), (FNumberOfItems + 10));
        // : Array [1..1000000,1..10000] of char;
        SetLength(lItemIDArray, (FNumberOfItems + 10));
        // : Array [1..10000] of string;
        SetLength(lExamineeIDArray, (FNumberOfExaminees + 10));
        // : Array [1..1000000] of string;
        SetLength(lItemInfoArray, (FNumberOfItems + 10), 10);
        // 5 Extra columns for multikeys
        SetLength(lItemScaleArray, (FNumberOfItems + 10), 3);
        SetLength(lRankingArray, (FNumberOfExaminees + 10), 3);
        { 1 = ranking, 2 = FGroup }
        SetLength(lReduceArray, (FNumberOfItems + 10), FNumCut + 1);
        SetLength(lSplitScore1, (FNumberOfExaminees + 2), 7);
        SetLength(lSplitScore2, (FNumberOfExaminees + 2), 7);
        SetLength(lScoreArray, (FNumberOfExaminees + 10), 11);
        SetLength(lNKeyArray, FNumberOfItems + 2, 2);

        if frmMain.difcheckbox.IsChecked then
        begin
          SetLength(lDifarray, succ(FNumberOfItems), 6 + StrToIntDef(frmMain.NumDifGroups.Text, 0));
          SetLength(lDifRankArray, succ(FNumberOfExaminees), 3);
          // 1 =total rank, 2=FGroup
        end;

        if frmMain.collusionbox.IsChecked then
        begin
          SetLength(FBBOindexArray, succ(FNumberOfExaminees), succ(FNumberOfExaminees));
          SetLength(FBBOFlags, succ(FNumberOfExaminees));
        end;

        { ---First, Read in FItem IDs to array--- }
        frmMain.ProgressLabel.Text := 'Reading in item control file';

        FNumberItemsSoFar := 0;
        FScoredItems := 0;
        FPreTest := 0;
        lFileRow := 0;
        lIDBegin := StrToIntDef(frmMain.IDColumnsBegin.Text, 0);

        if lIDBegin = 0 then
          lIDBegin := 1;

        lColBox := AItemResponsesBeginInColumn;
        lExamineeIDColumns := StrToIntDef(frmMain.eTab2Numbers1.Text, 0);
        lGroupCodeColumn := StrToIntDef(frmMain.GroupColumn.Text, 0);
        IsZero := False;
        // Number of Items built up.
        FNumDomains := 0;

        for i := 1 to FNumberOfItems do
          if lTypeOk then
          begin
            inc(FNumberItemsSoFar);

            { ------Control lColumn 1: FItem ID ------ }
            SetLength(lIDsArray, 100);
            lIDSpace := 0;
            repeat
              vv := 0;
              Read(lItemInfoFile, lChar);

              if i = 1 then
                vv := CharCode(lChar);
              // if vv < 100 then INC(lIDSpace);
              If (vv < 100) and (lChar <> lTabChar) then
                lIDsArray[lIDSpace] := lChar;
              If (lChar = lTabChar) and (lIDSpace = 1) then
                lIDsArray[lIDSpace] := ' ';
              if vv < 100 then
                inc(lIDSpace);
            until lChar = controldelim;

            SetLength(lIDsArray, pred(lIDSpace));
            lItemIDArray[FNumberItemsSoFar] := GetStrFromCharArr(lIDsArray);

            { ------Control lColumn 2: FKey ------ }
            Read(lItemInfoFile, lChar);

            if (lChar = controldelim) then
              lKeyOk := False;

            lItemInfoArray[FNumberItemsSoFar, 1] := UpCase(lChar); { FKey }
            Read(lItemInfoFile, lChar); { delimiter }
            lTempInt := 4;
            numkey := 1;

            while (lChar <> controldelim) do
            begin
              lItemInfoArray[FNumberItemsSoFar, lTempInt] := UpCase(lChar);
              // Multi FKey
              inc(lTempInt);
              inc(numkey);
              Read(lItemInfoFile, lChar); // next char
            end;

            Read(lItemInfoFile, lChar); // delimiter
            lNKeyArray[FNumberItemsSoFar, 1] := numkey;

            { ------Control lColumn 3: Number Alternatives ------ }
            k := 0;

            while lChar <> controldelim do // 16-1-2
            begin
              inc(k); // increments k by 1
              readformat[k] := lChar; // Multi FKey
              Read(lItemInfoFile, lChar);
            end;

            for j := succ(k) to 50 do
              readformat[j] := chr(0);
            // fills all remaining spots with Null char

            contint := PWideChar(@readformat);
            // Converts ReadFormat (array of Char) to string named ContInt
            if not TryStrToInt(contint, lTempInt) then
            begin
              lTempInt := 0;
              Log.d('Failed to convert "%s" to Integer', [contint]);
            end;


            if lTempInt > 15 then
              lAltOk := False
            else
              lItemScaleArray[i, 1] := lTempInt;

            { ------Control lColumn 4: Domain ------ }
            lChar := chr(0);
            lColumn := 0;
            lSameDomain := False;

            for lRow := 1 to 50 do
              lReadIDArray[lRow] := chr(0);

            while lChar <> controldelim do // meaningful domain labels!
            begin
              Read(lItemInfoFile, lChar);
              inc(lColumn);

              if lChar <> controldelim then
                lReadIDArray[lColumn] := lChar;
            end; // 18-1-2

            temp1 := PWideChar(@lReadIDArray);

            if FNumberItemsSoFar = 1 then
              FStoreDomain[1] := temp1;
            if FNumberItemsSoFar = 1 then
              lItemScaleArray[FNumberItemsSoFar, 2] := 1;

            if FNumberItemsSoFar = 1 then
              FNumDomains := 1
            else
              for j := 1 to FNumDomains do
                if FStoreDomain[j] = temp1 then
                  lSameDomain := True;

            if lSameDomain then
              for j := 1 to FNumDomains do
                if FStoreDomain[j] = temp1 then
                  lItemScaleArray[FNumberItemsSoFar, 2] := j;

            if (lSameDomain = False) and (FNumberItemsSoFar > 1) then
            begin
              inc(FNumDomains);
              FStoreDomain[FNumDomains] := temp1;
              lItemScaleArray[FNumberItemsSoFar, 2] := FNumDomains;
            end;

            for j := 1 to lColumn do
              lReadIDArray[j] := chr(0);

            { ------Control lColumn 5: Inclusion FCode ------ }
            Read(lItemInfoFile, lChar);
            lItemInfoArray[FNumberItemsSoFar, 2] := UpCase(lChar);
            { Scored }

            If UpperCase(lChar) = 'Y' then
              inc(FScoredItems);
            if UpperCase(lChar) = 'Y' then
              lAllPilot := False;
            if UpperCase(lChar) = 'P' then
              inc(FPreTest);

            Read(lItemInfoFile, lChar); // tab char  or delimiter

            { ------Control lColumn 6: FItem Type ------ }
            ReadLn(lItemInfoFile, lChar); // FItem type
            lItemInfoArray[FNumberItemsSoFar, 3] := UpCase(lChar);

            if UpperCase(lChar) <> 'P' then
              lAllPC := False;
            if (lItemInfoArray[FNumberItemsSoFar, 3] <> 'R') and (lItemInfoArray[FNumberItemsSoFar, 3] <> 'M') and
              (lItemInfoArray[FNumberItemsSoFar, 3] <> 'P') then
              lEndLine := False;

            if UpperCase(lChar) = 'P' then
              IsZero := True;
            if lItemInfoArray[FNumberItemsSoFar, 3] <> 'M' then
              FIsRate := True;
            if lItemInfoArray[FNumberItemsSoFar, 3] = 'M' then
              FIsMult := True;
            if (lItemInfoArray[FNumberItemsSoFar, 3] = 'M') and (lItemInfoArray[FNumberItemsSoFar, 2] = 'Y') then
              lScoredMult := True;
            if (lItemInfoArray[FNumberItemsSoFar, 3] <> 'M') and (lItemInfoArray[FNumberItemsSoFar, 2] = 'Y') then
              scoredrate := True;

            if lItemInfoArray[FNumberItemsSoFar, 2] = 'Y' then
            begin
              if lItemInfoArray[FNumberItemsSoFar, 3] <> 'M' then
                FNumRate[1] := FNumRate[1] + 1;
              if lItemInfoArray[FNumberItemsSoFar, 3] <> 'M' then
                FDomRate[lItemScaleArray[FNumberItemsSoFar, 2]] := FDomRate[lItemScaleArray[FNumberItemsSoFar, 2]] + 1;
              if lItemInfoArray[FNumberItemsSoFar, 3] = 'M' then
                FNumMult[1] := FNumMult[1] + 1;
              if lItemInfoArray[FNumberItemsSoFar, 3] = 'M' then
                FDomMult[lItemScaleArray[FNumberItemsSoFar, 2]] := FDomMult[lItemScaleArray[FNumberItemsSoFar, 2]] + 1;
            end;
            if lItemInfoArray[FNumberItemsSoFar, 2] <> 'N' then
            begin
              if lItemInfoArray[FNumberItemsSoFar, 3] <> 'M' then
                FNumRate[2] := FNumRate[2] + 1;
              if lItemInfoArray[FNumberItemsSoFar, 3] = 'M' then
                FNumMult[2] := FNumMult[2] + 1;
            end;
            if lItemInfoArray[FNumberItemsSoFar, 2] = 'P' then
            begin
              if lItemInfoArray[FNumberItemsSoFar, 3] <> 'M' then
                FNumRate[3] := FNumRate[3] + 1;
              if lItemInfoArray[FNumberItemsSoFar, 3] = 'M' then
                FNumMult[3] := FNumMult[3] + 1;
            end;
            if lItemInfoArray[FNumberItemsSoFar, 2] = 'N' then
            begin
              if lItemInfoArray[FNumberItemsSoFar, 3] <> 'M' then
                FNumRate[4] := FNumRate[4] + 1;
              if lItemInfoArray[FNumberItemsSoFar, 3] = 'M' then
                FNumMult[4] := FNumMult[4] + 1;
            end;
            if (lItemScaleArray[FNumberItemsSoFar, 1] = 2) or (lItemInfoArray[FNumberItemsSoFar, 3] = 'M') then
              lItemInfoArray[FNumberItemsSoFar, 9] := 'D'
            else
              lItemInfoArray[FNumberItemsSoFar, 9] := 'P';
          end;
        CloseFile(lItemInfoFile);
        // End of reading in the control file with 6 columns

        { -----error trapping # of items-------- }
        if not frmMain.delimitbox.IsChecked then
        begin
          for lRow := 1 to FNumberOfItems + pred(AItemResponsesBeginInColumn) do
          begin
            Read(lDataMatrixFile, lChar);

            if (lChar = chr(48)) and (lRow > AItemResponsesBeginInColumn) and
              (lItemInfoArray[lRow - succ(AItemResponsesBeginInColumn), 3] <> 'P') and (chr(48) <> FOmitChar) and
              (chr(48) <> FNAChar) then
              lZeroOk := False;

            if lZeroOk = False then
              FAllOk := False;
            if lChar = chr(13) then
              FAllOk := False;
          end;
        end; // not delimited input

        for lRow := 1 to FNumberOfItems do
          Read(lDataMatrixFile, lChar);

        Reset(lDataMatrixFile);

        // ����� ������������� ����� ������
        if not FAllOk and lColumnOk and lTypeOk and lZeroOk then
        begin
          frmMain.mLog.lines.Add('Please check the data file format specifications');
          Application.ProcessMessages;
        end;

        if mrf1 <> 'GUI' then
        begin
          if not FAllOk and lColumnOk and lTypeOk and lZeroOk then
            frmMain.mLog.lines.Add('ERROR! Check the data file format specifications');
          if not lTypeOk then
            frmMain.mLog.lines.Add('ERROR! Item control file does not include the item type');
          if not lZeroOk then
            frmMain.mLog.lines.Add('ERROR! Valid item responses of 0 were identified');
          if not FAllOk and lColumnOk and lTypeOk then
            frmMain.mLog.lines.Add('');
          Application.ProcessMessages;
        end;

        if not lKeyOk then
          frmMain.mLog.lines.Add('ERROR! The key column has missing values');
        if not lAltOk then
          frmMain.mLog.lines.Add('ERROR! The number of alternatives column includes a non-numeric character');
        if not lDomOk then
          frmMain.mLog.lines.Add('ERROR! The domain column includes a non-numeric character');
        if not lEndLine then
          frmMain.mLog.lines.Add('ERROR! Check the item type column, an unrecognized character was identified' + #13 +
            'Item type includes M = Multiple Choice, R = Rating Scale, P = Partial Credit');
        if not lZeroOk then
          frmMain.mLog.lines.Add('At least one valid item response of 0 was identified' + #13 +
            'The item control file specified that item responses began at 1 for that item');
        Application.ProcessMessages;
      end;

      frmMain.ProgressBar.Value := frmMain.ProgressBar.Value + 1;
      Application.ProcessMessages;

      if not lAltOk or not lDomOk or not lEndLine then
        FAllOk := False;

      { -------------------Reading in first line of ITAP Header---------------------- }
      if frmMain.cbDataMatrixFile.IsChecked then
      begin
        // Make change here.
        try
          Read(lDataMatrixFile, FNumberOfItems);
        except
          FAllOk := False;
        end;

        Read(lDataMatrixFile, lChar);
        Read(lDataMatrixFile, lChar);
        FOmitChar := UpCase(lChar);
        Read(lDataMatrixFile, lChar);
        Read(lDataMatrixFile, lChar);
        FNAChar := UpCase(lChar);

        try
          Read(lDataMatrixFile, lExamineeIDColumns);
        except
          FAllOk := False;
        end;

        ReadLn(lDataMatrixFile);

        if not FAllOk then
          ErrorMessage := ('Please select an input file with an ITEMAN 3 Header');
        if (mrf1 <> 'GUI') and not FAllOk then
        begin
          frmMain.mLog.lines.Add('ERROR! Data file does not have an ITEMAN 3 Header');
          frmMain.mLog.lines.Add('');
          Application.ProcessMessages;
        end;

        if FAllOk then
        begin
          { --- Demo Limitation // Set Number of Items if this is AppPath demo --- }
          if frmMain.Is_Demo then
            if FNumberOfItems > 100 then
              FNumberOfItems := 100;

          // Define all dynamic arrays - with some extra room just in case
          SetLength(lResponseArray, (FNumberOfExaminees + 10), (FNumberOfItems + 10));
          // : Array [1..1000000,1..10000] of char;
          SetLength(lItemIDArray, (FNumberOfItems + 10));
          // : Array [1..10000] of string;
          SetLength(lExamineeIDArray, (FNumberOfExaminees + 10));
          // : Array [1..1000000] of string;
          SetLength(lItemInfoArray, (FNumberOfItems + 10), 10);
          // Extra columns for multikeys
          SetLength(lItemScaleArray, (FNumberOfItems + 10), 3);
          SetLength(lRankingArray, (FNumberOfExaminees + 10), 3);
          { 1 = ranking, 2 = FGroup }
          SetLength(lReduceArray, (FNumberOfItems + 10), FNumCut + 1);
          SetLength(lSplitScore1, (FNumberOfExaminees + 2), 7);
          SetLength(lSplitScore2, (FNumberOfExaminees + 2), 7);
          SetLength(lScoreArray, (FNumberOfExaminees + 10), 11);
          SetLength(lNKeyArray, FNumberOfItems + 2, 2);

          if frmMain.difcheckbox.IsChecked then
            SetLength(lDifarray, FNumberOfItems + 1, 6 + StrToIntDef(frmMain.NumDifGroups.Text, 0));
          if frmMain.difcheckbox.IsChecked then
            SetLength(lDifRankArray, FNumberOfExaminees + 1, 3);
          // 1 =total rank, 2=FGroup

          FScoredItems := 0;
          FPreTest := 0;

          for lRow := 1 to FNumberOfItems do
          begin
            Read(lDataMatrixFile, lChar);
            lItemInfoArray[lRow, 1] := UpCase(lChar);

            if (lChar = '+') or (lChar = '-') then
              lItemInfoArray[lRow, 3] := 'R'
            else
              lItemInfoArray[lRow, 3] := 'M';

            if lItemInfoArray[lRow, 3] = 'R' then
              FIsRate := True;
            if lItemInfoArray[lRow, 3] = 'M' then
              FIsMult := True;
            lItemIDArray[lRow] := IntToStr(lRow);
            // setting lItemIDArray to FItem sequence number
          end;

          ReadLn(lDataMatrixFile); // finishing the line

          for lRow := 1 to FNumberOfItems do
          begin
            Read(lDataMatrixFile, lChar); // number of alternatives
            lTempInt := StrToInt(lChar);
            lItemScaleArray[lRow, 1] := lTempInt;
            lItemScaleArray[lRow, 2] := 1;
            FStoreDomain[1] := '1';
          end;

          ReadLn(lDataMatrixFile); // finishing off # alts

          for lRow := 1 to FNumberOfItems do
          begin
            Read(lDataMatrixFile, lChar);
            lNKeyArray[lRow, 1] := 1;

            if UpperCase(lChar) = 'P' then
              lItemInfoArray[lRow, 2] := 'P'
            else if UpperCase(lChar) = 'Y' then
              lItemInfoArray[lRow, 2] := 'Y'
            else if UpperCase(lChar) <> 'N' then
              lItemInfoArray[lRow, 2] := 'Y'
            else
              lItemInfoArray[lRow, 2] := 'N';

            if UpperCase(lChar) = 'Y' then
              FScoredItems := FScoredItems + 1;
            if UpperCase(lChar) = 'Y' then
              lAllPilot := False;
            if UpperCase(lChar) = 'P' then
              FPreTest := FPreTest + 1;

            if lItemInfoArray[lRow, 2] = 'Y' then
            begin
              if lItemInfoArray[lRow, 3] <> 'M' then
                FNumRate[1] := FNumRate[1] + 1;
              if lItemInfoArray[lRow, 3] = 'M' then
                FNumMult[1] := FNumMult[1] + 1;
              if lItemInfoArray[lRow, 3] <> 'M' then
                FDomRate[1] := FDomRate[1] + 1;
              if lItemInfoArray[lRow, 3] = 'M' then
                FDomMult[1] := FDomMult[1] + 1;
            end;

            if lItemInfoArray[lRow, 2] <> 'N' then
            begin
              if lItemInfoArray[lRow, 3] <> 'M' then
                FNumRate[2] := FNumRate[2] + 1;
              if lItemInfoArray[lRow, 3] = 'M' then
                FNumMult[2] := FNumMult[2] + 1;
            end;

            if lItemInfoArray[lRow, 2] = 'P' then
            begin
              if lItemInfoArray[lRow, 3] <> 'M' then
                FNumRate[3] := FNumRate[3] + 1;
              if lItemInfoArray[lRow, 3] = 'M' then
                FNumMult[3] := FNumMult[3] + 1;
            end;

            if lItemInfoArray[lRow, 2] = 'N' then
            begin
              if lItemInfoArray[lRow, 3] <> 'M' then
                FNumRate[4] := FNumRate[4] + 1;
              if lItemInfoArray[lRow, 3] = 'M' then
                FNumMult[4] := FNumMult[4] + 1;
            end;

            if (lItemInfoArray[lRow, 3] = 'M') and (lItemInfoArray[lRow, 2] = 'Y') then
              lScoredMult := True;
            if (lItemInfoArray[lRow, 3] <> 'M') and (lItemInfoArray[lRow, 2] = 'Y') then
              scoredrate := True;

            if (lItemScaleArray[lRow, 1] = 2) or (lItemInfoArray[lRow, 3] = 'M') then
              lItemInfoArray[lRow, 9] := 'D'
            else
              lItemInfoArray[lRow, 9] := 'P'; // needed for DIF analysis
          end;

          ReadLn(lDataMatrixFile); // finishing 'er off

          for lRow := 1 to FNumberOfItems + lExamineeIDColumns do
          begin
            Read(lDataMatrixFile, lChar);

            if (lChar = chr(48)) and (chr(48) <> FOmitChar) and (chr(48) <> FNAChar) and (lRow > lExamineeIDColumns)
            then
            begin
              lZeroOk := False;
              FAllOk := False;
            end;

            if lChar = chr(13) then
              FAllOk := False;
          end;

          if not FAllOk and lAllPilot and frmMain.cbDataMatrixFile.IsChecked then
            frmMain.mLog.lines.Add('There were no scored items specified in the Iteman 3 Header. ' + #13 +
              'Please check the Header and make sure scored items are coded as Y');
          if not FAllOk and lZeroOk and not lAllPilot then
            frmMain.mLog.lines.Add
              ('Please check the number of items or number of ID columns specified in the ITEMAN 3 Header');
          if not FAllOk and not lZeroOk then
            frmMain.mLog.lines.Add
              ('Valid item responses of 0 were found. The ITEMAN 3 Header does not support item responses that begin at 0.');

          if mrf1 <> 'GUI' then
          begin
            if not FAllOk then
              frmMain.mLog.lines.Add('ERROR! Check the number of items/ID columns in the ITEMAN 3 Header');
            if not lZeroOk then
              frmMain.mLog.lines.Add('ERROR! Item responses cannot begin at 0 when an ITEMAN 3 Header is used');
            if not FAllOk then
              frmMain.mLog.lines.Add('');
            if not FAllOk and lAllPilot and frmMain.cbDataMatrixFile.IsChecked then
              frmMain.mLog.lines.Add('There were no scored items specified in the Iteman 3 Header');
          end;
          Application.ProcessMessages;

          lIDBegin := 1;
          lColBox := succ(lExamineeIDColumns);
          Reset(lDataMatrixFile);
          // resetting matrix so it begins with first FItem

          for lRow := 1 to 4 do
            ReadLn(lDataMatrixFile);
        end; // if FAllOk = True
      end; // if cbDataMatrixFile.IsChecked

      FValidItems := FPreTest + FScoredItems;

      if not frmMain.cbDataMatrixFile.IsChecked then
        lExamineeIDColumns := StrToIntDef(frmMain.eTab2Numbers1.Text, 0);

      if lAllPilot then
        FAllOk := False;
      if not FAllOk and lAllPilot and not frmMain.cbDataMatrixFile.IsChecked then
        frmMain.mLog.lines.Add('There were no scored items specified in the item control file. ' + #13 +
          'Please check the control file and make sure scored items are coded as Y');
      if (mrf1 <> 'GUI') and not FAllOk and lAllPilot and not frmMain.cbDataMatrixFile.IsChecked then
        frmMain.mLog.lines.Add('There were no scored items specified in the item control file');
      Application.ProcessMessages;
      if FAllOk then
      begin
        { ----------------------Read in FItem responses to array----------------------------- }
        frmMain.ProgressLabel.Text := 'Reading in data file';
        i := 0; // examinee without all FItem responses, if that occurs
        lMaxExamId := 0;
        // maximum examinee id chars for delimited input, for correct output file

        for lRow := 1 to FNumberOfExaminees do
        begin // 18
          // avc := 0;

          if not frmMain.delimitbox.IsChecked then
          begin
            if lExamineeIDColumns > 0 then
            begin // 18-1 error trapping because lIDBegin will go negative!
              { Read nothing until ID starts - part of being dynamic }
              if lGroupCodeColumn < lIDBegin then
              // if FGroup FCode appears BEFORE ID, parse it out, then Read to begin of ID
              begin
                for lColumn := 1 to pred(lGroupCodeColumn) do
                  Read(lDataMatrixFile, lChar);

                if frmMain.difcheckbox.IsChecked then
                  Read(lDataMatrixFile, lChar);
                if frmMain.difcheckbox.IsChecked then
                  lResponseArray[lRow, 0] := UpCase(lChar); // DIF FCode
                for lColumn := 1 to lIDBegin - lGroupCodeColumn - 1 do
                  Read(lDataMatrixFile, lChar);
                // getting to begin of Real ID

                For lColumn := 1 to lExamineeIDColumns do
                begin
                  Read(lDataMatrixFile, lChar);
                  lReadIDArray[lColumn] := lChar;
                end;

                lExamineeID := PWideChar(@lReadIDArray);
                lExamineeIDArray[lRow] := lExamineeID;

                for lColumn := 1 to lColBox - lIDBegin - lExamineeIDColumns do
                  Read(lDataMatrixFile, lChar);
              end
              else
                for lColumn := 1 to lIDBegin - 1 do
                  Read(lDataMatrixFile, lChar);

              if (lGroupCodeColumn > lIDBegin) and (lGroupCodeColumn < (lIDBegin + lExamineeIDColumns)) then
              // if difcode is embedded in examinee ID, parse it out!
              begin
                idread := 0;

                for lColumn := 1 to lGroupCodeColumn - lIDBegin do
                begin
                  inc(idread);
                  Read(lDataMatrixFile, lChar);
                  lReadIDArray[idread] := lChar;
                  { need to Reset to all nulls? }
                end;

                if frmMain.difcheckbox.IsChecked then
                  Read(lDataMatrixFile, lChar);
                if frmMain.difcheckbox.IsChecked then
                  lResponseArray[lRow, 0] := UpCase(lChar);

                for lColumn := 1 to (lExamineeIDColumns + lIDBegin - 1) - lGroupCodeColumn do
                begin
                  inc(idread);
                  Read(lDataMatrixFile, lChar);
                  lReadIDArray[idread] := lChar;
                end;

                lExamineeID := PWideChar(@lReadIDArray);
                lExamineeIDArray[lRow] := lExamineeID;

                for lColumn := 1 to (lColBox - lIDBegin - lExamineeIDColumns) do
                  Read(lDataMatrixFile, lChar);
                // getting to correct Data lColumn
              end;

              if (lGroupCodeColumn > lIDBegin) and (lGroupCodeColumn > (lIDBegin + lExamineeIDColumns - 1)) then
              // if difcode appears after examinee ID, Read in ID first
              begin
                { Read ID data }
                for lColumn := 1 to lExamineeIDColumns do // 18-1-2
                begin
                  Read(lDataMatrixFile, lChar);
                  lReadIDArray[lColumn] := lChar;
                  { need to Reset to all nulls? }
                end; // 18-1-2

                lExamineeID := PWideChar(@lReadIDArray);
                lExamineeIDArray[lRow] := lExamineeID;

                for lColumn := 1 to (lGroupCodeColumn - lExamineeIDColumns - lIDBegin) do
                  Read(lDataMatrixFile, lChar); // getting to DIF FCode

                if frmMain.difcheckbox.IsChecked then
                  Read(lDataMatrixFile, lChar);
                if frmMain.difcheckbox.IsChecked then
                  lResponseArray[lRow, 0] := UpCase(lChar); // DIF FCode

                for lColumn := 1 to (lColBox - lGroupCodeColumn - 1) do
                  Read(lDataMatrixFile, lChar);
                // getting to correct Data lColumn
              end; // if difcode appears after ID characters
            end; // if lExamineeIDColumns > 0 /18-1

            if frmMain.difcheckbox.IsChecked then
              if lResponseArray[lRow, 0] = UpCase(frmMain.group1code.Text[1]) then
                lResponseArray[lRow, 0] := '1'
              else if lResponseArray[lRow, 0] = UpCase(frmMain.group2code.Text[1]) then
                lResponseArray[lRow, 0] := '2'
              else
                lResponseArray[lRow, 0] := '0';
          end; // if delimitbox.IsChecked = False statement

          if frmMain.delimitbox.IsChecked then
          begin
            for j := 1 to 200 do
              lReadIDArray[j] := chr(0);

            { Read ID data }
            if frmMain.commabox.IsChecked then
              delim1 := ','
            else
              delim1 := chr(9);

            if frmMain.includeidbox.IsChecked then
            begin
              lColumn := 0;
              lChar := chr(0);

              while lChar <> delim1 do // 18-1-2
              begin
                Read(lDataMatrixFile, lChar);
                inc(lColumn);

                if lChar <> delim1 then
                  lReadIDArray[lColumn] := lChar;
              end; // 18-1-2

              if lColumn > lMaxExamId then
                lMaxExamId := lColumn;
              lExamineeID := PWideChar(@lReadIDArray);
              lExamineeIDArray[lRow] := lExamineeID;
            end
            else
            begin
              lExamineeIDArray[lRow] := IntToStr(lRow);
              // so examinee ID isn't null! and equals row number
              x := power(10, lMaxExamId);

              if 0 = (lRow mod round(x)) then
                inc(lMaxExamId);
              // this increments maximum # characters for examinee ID
            end;

            if frmMain.difcheckbox.IsChecked then
            begin
              Read(lDataMatrixFile, lChar);

              if UpperCase(lChar) = UpCase(frmMain.group1code.Text[1]) then
                lResponseArray[lRow, 0] := '1'
              else if UpperCase(lChar) = UpCase(frmMain.group2code.Text[1]) then
                lResponseArray[lRow, 0] := '2'
              else
                lResponseArray[lRow, 0] := '0';

              Read(lDataMatrixFile, lChar);
            end; // reading in dif FCode.
          end; // if delimitbox.IsChecked statement

          lastvalid := FNumberOfItems;
          errno := 0;

          { Read and put the response in the array }
          for lColumn := 1 to FNumberOfItems do
          begin
            if not frmMain.delimitbox.IsChecked then
              Read(lDataMatrixFile, lChar);

            if frmMain.delimitbox.IsChecked then
            begin // trapping null responses
              if lColumn = 1 then
                Read(lDataMatrixFile, lChar)
              else if not lIsDelim then
                Read(lDataMatrixFile, lChar, lChar)
              else
                Read(lDataMatrixFile, lChar);

              if lChar = delim1 then
                lIsDelim := True
              else
                lIsDelim := False;
              if lIsDelim then
                lChar := UpCase(FNAChar);
            end;

            if (lChar = chr(13)) or (lChar = chr(26)) then
            begin
              lChar := FNAChar;
              lResponseArray[lRow, lColumn] := UpCase(FNAChar);
              errno := lRow;
              // avc:= lRow;
              lastvalid := lColumn;
            end;

            if errno <> 0 then
              i := errno;
            if errno <> 0 then
              for j := lastvalid to FNumberOfItems do
                lResponseArray[lRow, j] := UpCase(FNAChar);

            lResponseArray[lRow, lColumn] := UpCase(lChar);

            if errno <> 0 then
              break;

            // checking for unrecognized characters
            if UpCase(lChar) <> UpperCase(FOmitChar) then
              if lChar = chr(32) then
                lChar := UpCase(FNAChar); // converting spaces to NA's

            if (UpCase(lChar) <> UpperCase(FOmitChar)) and (UpperCase(lChar) <> UpperCase(FNAChar)) then
            begin
              vv := ConvertResponseToInteger(UpCase(lChar), lItemInfoArray, lColumn, lItemScaleArray[lColumn, 1]);

              if vv = succ(lItemScaleArray[lColumn, 1]) then
                lResponseArray[lRow, lColumn] := UpCase(FOmitChar);
              // setting unrecognized chars to omit
            end; // begin-end statement added to fix N to O conversion bug (Feb.2013)

            if (UpperCase(lChar) <> UpperCase(FOmitChar)) and (UpperCase(lChar) <> UpperCase(FNAChar)) then
            begin
              datacontrolerror[lColumn] := (IsNumeric(UpCase(lChar)) <> IsNumeric(lItemInfoArray[lColumn, 1])) and
                (IsAlphaBetic(UpCase(lChar)) <> IsAlphaBetic(lItemInfoArray[lColumn, 1]));
              if datacontrolerror[lColumn] then
                lMismatch := True;
            end;

            lNumberOfEntries := succ(lNumberOfEntries);
          end; // for lColumn loop (FNumberOfItems)

          if errno = 0 then
            ReadLn(lDataMatrixFile);
          if errno <> 0 then
          begin
            Reset(lDataMatrixFile);
            if frmMain.cbDataMatrixFile.IsChecked then
              k := 4
            else
              k := 0;

            for j := 1 to lRow + k do
              ReadLn(lDataMatrixFile);
          end;
        end; { for lRow loop }

        matchwarning := '';
        revmatchwarning := '';

        for lColumn := 1 to FNumberOfItems do
          if datacontrolerror[lColumn] = True then
          begin
            if (IsNumeric(lItemInfoArray[lColumn, 1])) and (matchwarning = '') then
              matchwarning := 'numeric';
            if (IsAlphaBetic(lItemInfoArray[lColumn, 1])) and (matchwarning = '') then
              matchwarning := 'alphabetical';
            if (IsNumeric(lItemInfoArray[lColumn, 1])) and (matchwarning[1] = 'a') then
              matchwarning := 'alphabetical or numeric';
            if (IsAlphaBetic(lItemInfoArray[lColumn, 1])) and (matchwarning[1] = 'n') then
              matchwarning := 'numeric or alphabetical';
            if (IsNumeric(lItemInfoArray[lColumn, 1])) and (revmatchwarning = '') then
              revmatchwarning := 'alphabetical';
            if (IsAlphaBetic(lItemInfoArray[lColumn, 1])) and (revmatchwarning = '') then
              revmatchwarning := 'numeric';
            if (IsNumeric(lItemInfoArray[lColumn, 1])) and (revmatchwarning[1] = 'n') then
              revmatchwarning := 'numeric or alphabetical';
            if (IsAlphaBetic(lItemInfoArray[lColumn, 1])) and (revmatchwarning[1] = 'a') then
              revmatchwarning := 'alphabetical or numeric';

            for lRow := 1 to FNumberOfExaminees do
            begin
              lChar := lResponseArray[lRow, lColumn];

              if (lChar <> FOmitChar) and (lChar <> FNAChar) then
              begin
                if (IsAlphaBetic(lItemInfoArray[lColumn, 1])) and not IsAlphaBetic(lChar) then
                  lResponseArray[lRow, lColumn] := chr(64 + RespToInt(lChar, lItemInfoArray, lColumn, lItemScaleArray));
                if IsNumeric(lItemInfoArray[lColumn, 1]) and not IsNumeric(lChar) then
                  lResponseArray[lRow, lColumn] := chr(48 + RespToInt(lChar, lItemInfoArray, lColumn, lItemScaleArray));
              end; // if not omit or NotAdmin
            end; // lRow loop
          end;

        if lMismatch then
          frmMain.mLog.lines.Add('The item control file specified that all item responses were ' + matchwarning + #13 +
            'but the item responses included ' + revmatchwarning + ' characters. The mismatched item responses ' + #13 +
            'were converted to be ' + matchwarning + ' for those item(s).');

        if i <> 0 then
          frmMain.mLog.lines.Add('Check the data matrix file, examinee ' + IntToStr(i) + ' did not respond to all ' +
            IntToStr(FNumberOfItems) + ' items.');

        frmMain.ProgressBar.Value := frmMain.ProgressBar.Value + 1;
        Application.ProcessMessages;

        { Find the number of domains }
        i := 0;
        for lRow := 1 to FNumberOfItems do
          if (lItemScaleArray[lRow, 2] > i) and (lItemInfoArray[lRow, 2] = 'Y') then
            i := lItemScaleArray[lRow, 2];

        FNumDomains := i;
        i := 0; // resetting to 0 to avoid False error message~~~

        { Define Progress Bar }
        frmMain.ProgressBar.Max := 0;
        frmMain.ProgressBar.Max := 5 * FNumberOfItems;

        if FIsRate then
          frmMain.ProgressBar.Max := frmMain.ProgressBar.Max + FNumberOfItems;
        if FIsMult then
          frmMain.ProgressBar.Max := frmMain.ProgressBar.Max + FNumberOfItems;
        if FNumDomains > 1 then
          frmMain.ProgressBar.Max := frmMain.ProgressBar.Max + 2 * FNumberOfItems * FNumDomains;

        frmMain.ProgressBar.Value := 0;

        if not FRespOK then
          frmMain.mLog.lines.Add('At least one unidentified response character was identified' + #13 +
            ' and will be scored as incorrect.');

        Application.ProcessMessages;

        { ---Read external scores to array if selected--- }
        If frmMain.cbExternalScoreFile.IsChecked then
        begin
          lFileRow := 0;

          while not EOF(lExternalScoreFile) do
          begin
            ReadLn(lExternalScoreFile, lExternalScore);
            lFileRow := succ(lFileRow);
            FExternalScoreArray[lFileRow] := lExternalScore;
          end;
        end;

        frmMain.ProgressBar.Value := frmMain.ProgressBar.Value + 1;
        SetLength(lDomScaleArray, FNumberOfExaminees + 3, FNumDomains + 2);
        Application.ProcessMessages;

        { ------Find the max number of options------ }
        FMaxItemOptions := 0;
        for lRow := 1 to FNumberOfItems do
          if lItemInfoArray[lRow, 2] <> 'N' then
            if lItemScaleArray[lRow, 1] > FMaxItemOptions then
              FMaxItemOptions := lItemScaleArray[lRow, 1];
        { high, med, low for each option }

        { ----------Reset lScoreArray to null!------------------ }
        for lRow := 1 to FNumberOfExaminees do
          for lColumn := 1 to 10 do
            lScoreArray[lRow, lColumn] := 0;

        { -------------Assigning the control file if ITEMAN HEADER------------------ }
        if frmMain.SaveControl.IsChecked then
        begin
          AssignFile(lControlText, ChangeFileExt(outp, '') + ' Control.txt');
          ReWrite(lControlText);
          i := 0;
        end;

        { --------------Score all items------------- }
        frmMain.ProgressLabel.Text := 'Scoring Responses';

        for lRow := 1 to FNumberOfExaminees do
        begin
          if frmMain.ScoredMatrixBox.IsChecked then
            write(lScoredMatrixFile, lExamineeIDArray[lRow]);

          for lColumn := 1 to FNumberOfItems do
          begin
            inc(i);

            if (lRow = 1) and (frmMain.SaveControl.IsChecked) and (lItemInfoArray[lColumn, 3] = 'M') then
            begin
              write(lControlText, 'Item ' + IntToStr(i));

              if frmMain.ScoredMatrixBox.IsChecked then
                write(lControlText, chr(9) + '1' + chr(9) + '2' + chr(9) + '1' + chr(9) + lItemInfoArray[lColumn, 2])
              else
                write(lControlText, chr(9) + lItemInfoArray[lColumn, 1] + chr(9) + IntToStr(lItemScaleArray[lColumn, 1])
                  + chr(9) + IntToStr(lItemScaleArray[lColumn, 2]) + chr(9) + lItemInfoArray[lColumn, 2]);

              if frmMain.ScoredMatrixBox.IsChecked then
                writeln(lControlText, chr(9) + 'P');
              if not frmMain.ScoredMatrixBox.IsChecked then
                writeln(lControlText, chr(9) + 'M');
            end; // if cbDataMatrixFile.IsChecked

            if (lRow = 1) and (frmMain.SaveControl.IsChecked) and (lItemInfoArray[lColumn, 3] = 'R') then
            begin
              write(lControlText, 'Item ' + IntToStr(i));
              write(lControlText, chr(9) + '+' + chr(9) + IntToStr(lItemScaleArray[lColumn, 1]) + chr(9) +
                IntToStr(lItemScaleArray[lColumn, 2]) + chr(9) + lItemInfoArray[lColumn, 2]);
              writeln(lControlText, chr(9) + 'R');
            end; // if cbDataMatrixFile.IsChecked
            if (lRow = 1) and (frmMain.SaveControl.IsChecked) and (lItemInfoArray[lColumn, 3] = 'P') then
            begin
              write(lControlText, 'Item ' + IntToStr(i));
              write(lControlText, chr(9) + '+' + chr(9) + IntToStr(lItemScaleArray[lColumn, 1]) + chr(9) +
                IntToStr(lItemScaleArray[lColumn, 2]) + chr(9) + lItemInfoArray[lColumn, 2]);
              writeln(lControlText, chr(9) + 'P');
            end; // if cbDataMatrixFile.IsChecked

            if lResponseArray[lRow, lColumn] <> FOmitChar then
              if lResponseArray[lRow, lColumn] <> FNAChar then
              begin
                if lItemInfoArray[lColumn, 3] = 'M' then
                  for j := 1 to lNKeyArray[lColumn, 1] do
                  begin
                    if j = 1 then
                      lTempInt := 0
                    else
                      lTempInt := 2;

                    if lResponseArray[lRow, lColumn] = lItemInfoArray[lColumn, lTempInt + j] then
                    begin { Correct }
                      if frmMain.ScoredMatrixBox.IsChecked then
                        write(lScoredMatrixFile, '1');
                      if lItemInfoArray[lColumn, 2] <> 'N' then
                        lScoreArray[lRow, 4] := lScoreArray[lRow, 4] + 1;
                      { all-FItem score }

                      if lItemInfoArray[lColumn, 2] = 'Y' then { FItem is scored }
                        lScoreArray[lRow, 1] := lScoreArray[lRow, 1] + 1;
                      { scored-FItem score }

                      if lItemInfoArray[lColumn, 2] = 'P' then { FItem is FPreTest }
                        lScoreArray[lRow, 5] := lScoreArray[lRow, 5] + 1;
                      { FPreTest-FItem score }
                    end
                    else if frmMain.ScoredMatrixBox.IsChecked then
                      write(lScoredMatrixFile, '0');
                  end; // for j to lNKeyArray

                if lItemInfoArray[lColumn, 3] <> 'M' then
                begin
                  if lItemInfoArray[lColumn, 2] <> 'N' then
                    lScoreArray[lRow, 4] := lScoreArray[lRow, 4] +
                      ConvertResponseToInteger(lResponseArray[lRow, lColumn], lItemInfoArray, lColumn,
                      lItemScaleArray[lColumn, 1]);
                  if lItemInfoArray[lColumn, 2] = 'Y' then
                    lScoreArray[lRow, 1] := lScoreArray[lRow, 1] +
                      ConvertResponseToInteger(lResponseArray[lRow, lColumn], lItemInfoArray, lColumn,
                      lItemScaleArray[lColumn, 1]);
                  if lItemInfoArray[lColumn, 2] = 'P' then
                    lScoreArray[lRow, 5] := lScoreArray[lRow, 5] +
                      ConvertResponseToInteger(lResponseArray[lRow, lColumn], lItemInfoArray, lColumn,
                      lItemScaleArray[lColumn, 1]);
                  if frmMain.ScoredMatrixBox.IsChecked then
                    write(lScoredMatrixFile, ConvertResponseToInteger(lResponseArray[lRow, lColumn], lItemInfoArray,
                      lColumn, lItemScaleArray[lColumn, 1]));
                end; // if lItemInfoArray = 'R' or rating scale
              end; // if <> FNAChar

            if frmMain.ScoredMatrixBox.IsChecked and frmMain.InclOmit.IsChecked then
              if lResponseArray[lRow, lColumn] = FOmitChar then
                write(lScoredMatrixFile, FOmitChar);

            if frmMain.ScoredMatrixBox.IsChecked and frmMain.InclNR.IsChecked then
              if lResponseArray[lRow, lColumn] = FNAChar then
                write(lScoredMatrixFile, FNAChar);

            if (frmMain.ScoredMatrixBox.IsChecked) and (frmMain.InclOmit.IsChecked = False) and
              (lItemInfoArray[lColumn, 3] = 'M') then
              if lResponseArray[lRow, lColumn] = FOmitChar then
                write(lScoredMatrixFile, '0');

            if (frmMain.ScoredMatrixBox.IsChecked) and (frmMain.InclNR.IsChecked = False) and
              (lItemInfoArray[lColumn, 3] = 'M') then
              if lResponseArray[lRow, lColumn] = FNAChar then
                write(lScoredMatrixFile, '0');
          end; // for lColumn to FNumberOfItems

          if frmMain.ScoredMatrixBox.IsChecked then
            writeln(lScoredMatrixFile);
        end;

        { ------------Call all the analysis procedures---------------------- }
        frmMain.ProgressBar.Value := frmMain.ProgressBar.Value + 10;
        frmMain.ProgressLabel.Text := 'Calculating statistics';
        Application.ProcessMessages;

        ScoreGroups(lRankingArray, lReduceArray, lScoreArray, lItemInfoArray, lResponseArray);
        // split sample into scoregroups
        SumStats(lItemInfoArray, lScoreArray);
        // compute mean,sd,var,min,and max for all scales
        OptionStats(lItemInfoArray, lResponseArray, lItemScaleArray, lNKeyArray, lScoreArray);
        // option rpbis, rbis, means , FOmch defined here
        ScaleStats(lItemInfoArray); // alpha, mean FP, and mean rpbis
        EstimateKR20Without(lResponseArray, lItemInfoArray, lItemScaleArray, lNKeyArray, lScoreArray);
        DomainStats(lItemInfoArray, lResponseArray, lItemScaleArray, lDomScaleArray, lNKeyArray);

        if FNumDomains > 1 then
          SetLength(lSplitDomain1, FNumberOfExaminees + 2, 1 + FNumDomains * 3);
        if FNumDomains > 1 then
          SetLength(lSplitDomain2, FNumberOfExaminees + 2, 1 + FNumDomains * 3);

        Reliability(lSplitScore1, lSplitScore2, lSplitDomain1, lSplitDomain2, lItemScaleArray, lItemInfoArray,
          lResponseArray);
        SetLength(lCsemArray, round(FTestStatsArray[1, 1]) + 2, 3);
        CSem(lItemScaleArray, lItemInfoArray, lResponseArray, lCsemArray);

        if (FNumberOfExaminees > 1) and (FTestStatsArray[1, 6] <= 1) then
        begin
          FTestStatsArray[1, 9] := FTestStatsArray[1, 3] * sqrt(1 - FTestStatsArray[1, 6]);
          if (FPreTest > 0) and (FTestStatsArray[2, 6] <= 1) then
            FTestStatsArray[2, 9] := FTestStatsArray[2, 3] * sqrt(1 - FTestStatsArray[2, 6]);
          if (FPreTest > 0) and (FTestStatsArray[3, 6] <= 1) then
            FTestStatsArray[3, 9] := FTestStatsArray[3, 3] * sqrt(1 - FTestStatsArray[3, 6]);

          if FNumDomains > 1 then
            for i := 1 to FNumDomains do
              if FDomainStatsArray[i, 6] <= 1 then
                FDomainStatsArray[i, 9] := FDomainStatsArray[i, 3] * sqrt(1 - FDomainStatsArray[i, 6]);
        end;

        if frmMain.difcheckbox.IsChecked then
          DIFgroups(lDifRankArray, lScoreArray, lItemInfoArray, lResponseArray);

        if frmMain.difcheckbox.IsChecked then
          if FNumberForDif > 0 then
          begin
            for lItem := 1 to FNumberOfItems do
              if (lItemInfoArray[lItem, 2] = 'Y') and (lItemInfoArray[lItem, 9] = 'D') then
                if frmMain.difcheckbox.IsChecked then
                  mantelH(lItem, lResponseArray, lItemInfoArray, lDifRankArray, lDifarray);
          end
          else
          begin
            frmMain.mLog.lines.Add('No recognized DIF character codes were found. DIF analysis could not be performed');
            Application.ProcessMessages;
            frmMain.difcheckbox.IsChecked := False;
          end;

        if frmMain.difcheckbox.IsChecked then
          for lRow := 1 to FNumberOfItems do
            if (lItemInfoArray[lRow, 2] = 'Y') and (lItemInfoArray[lRow, 9] = 'D') then
            begin
              if lDifarray[lRow, 4] > 0 then
                lDifarray[lRow, 5] := 1 - lDifarray[lRow, 5];

              lDifarray[lRow, 5] := lDifarray[lRow, 5] * 2;
            end;

        { -------------------Doing the raw frequency table 10-20-11---------------------- }
        for lRow := 1 to FNumberOfExaminees do
          lScoreArray[lRow, 0] := lScoreArray[lRow, 1];

        SetLength(lRawFrequency, round(FTestStatsArray[1, 5] - FTestStatsArray[1, 4]) + 3, 4);
        ShellSortInt(lScoreArray, FNumberOfExaminees);
        RawFreq(lRawFrequency, lScoreArray);

        { -------------------Doing BBOindex---------------------------------- }
        if frmMain.collusionbox.IsChecked then
          BBOindex(lItemInfoArray, lResponseArray, lItemScaleArray);

        { _________________   CSV OUTPUT SECTION   _________________ }
        CSVOutputSection();

        if frmMain.collusionbox.IsChecked then
        begin
          { _________________   BBO matrix output section  _________________ }
          Result := BBOMatrixOutputFile();
          if not Result then
            Exit;

          { _________________   DOCX OUTPUT SECTION   _________________ }
          // DOCXResultFileCreating();
        end;
      end; // ALLOK = TRUE
      { Close files etc. }
      frmMain.ProgressBar.Value := frmMain.ProgressBar.Value + 5;
      Application.ProcessMessages;

      if FNumberOfItems < itemcap then
      begin
        if UseMRF then
          OutputDir := path
        else
          OutputDir := path;

        if UseMRF then
          frmMain.docxoutput := ChangeFileExt(outp, '.docx')
        else
          frmMain.docxoutput := outp;

        // ���������� ���������� ������� � DOC ����
        try
          /// /          FExport.SaveToFile(frmMain.docxoutput);
          DOCXResultFileCreating();
        except
          on E: Exception do
          begin
            /// /          ErrorMessage:= E.Message;
            /// /          Result:= False;
            /// /          Exit;
          end;
        end;
      end;

      frmMain.ProgressBar.Value := frmMain.ProgressBar.Value + 5;
      frmMain.ProgressBar.Value := frmMain.ProgressBar.Value + 5;
      Application.ProcessMessages;

      if frmMain.ScoredMatrixBox.IsChecked then
        CloseFile(lScoredMatrixFile);

      frmMain.ProgressBar.Value := frmMain.ProgressBar.Value + 45;
      Application.ProcessMessages;

      if not UseMRF then
        lOutputFileName := '"' + frmMain.docxoutput + '"';
      if UseMRF then
        lOutputFileName := '"' + mrfname + '"';

      frmMain.ProgressBar.Value := 0;
      CloseFile(lDataMatrixFile);
      Application.ProcessMessages;

      if frmMain.SaveControl.IsChecked then
        CloseFile(lControlText);

      CloseFile(lOutputFile);
      CloseFile(lScoresFile);
      finalize(lResponseArray);
      finalize(lItemIDArray);
      finalize(lExamineeIDArray);
      finalize(lItemInfoArray);
      finalize(lItemScaleArray);
      finalize(lRankingArray);
      finalize(lSplitScore1);
      finalize(lSplitScore2);
      finalize(lCsemArray);
      finalize(lScoreArray);

      { if FAllOk and not usemrf then
        frmMain.DialogCall(); }
      // �������� � ���������� ���������� ������ ������
    end;

    Result := FAllOk;

    if Result then
      ErrorMessage := 'The program has completed the item analysis.';
  finally
    frmMain.ProgressLabel.Text := '';
  end;
end;

function TItemanCalculation.CharCode(aChar: Char): Integer;
var
  i: Integer;
begin
  Result := -1;
  for i := 0 to 200 do
    if (chr(i) = UpperCase(aChar)) then
      Result := i;
end;

function TItemanCalculation.GetStrFromCharArr(aCharsArray: TChar1Array): string;
var
  ifor: Integer;
begin
  Result := '';
  for ifor := 0 to pred(Length(aCharsArray)) do
    Result := Result + aCharsArray[ifor];
end;

function TItemanCalculation.GetTableSpec: TTableSpec;
begin
  SetLength(Result, 22, 2);

  with frmMain do
  begin
    Result[0, 0].Name := 'Number of examinees';
    Result[0, 0].Value := IntToStr(FNumberOfExaminees);
    Result[0, 1].Name := 'Total Items';
    Result[0, 1].Value := IntToStr(FNumberOfItems);
    Result[1, 0].Name := 'Scored Items';
    Result[1, 0].Value := IntToStr(FScoredItems);
    Result[1, 1].Name := 'Pretest Items';
    Result[1, 1].Value := formatfloat('###', FTestStatsArray[3, 1]);
    Result[2, 0].Name := 'Multiple Choice Items';
    Result[2, 0].Value := IntToStr(FNumMult[2]);
    Result[2, 1].Name := 'Polytomous Items';
    Result[2, 1].Value := IntToStr(FNumRate[2]);
    Result[3, 0].Name := 'Number of Domains';
    Result[3, 0].Value := IntToStr(FNumDomains);
    Result[3, 1].Name := 'External scores';
    // Result[3, 1].Value:= BoolToStr(extscorecheck.checked, 'Yes', 'No');
    Result[3, 1].Value := BoolToStr(cbExternalScoreFile.IsChecked, True);
    Result[4, 0].Name := 'Minimum P';
    Result[4, 0].Value := formatfloat('0.00', FMinP);
    Result[4, 1].Name := 'Maximum P';
    Result[4, 1].Value := formatfloat('0.00', FMaxP);
    Result[5, 0].Name := 'Minimum item mean';
    // Result[5, 0].Value:= FormatFloat('0.00',MainForm.minmeanbox.value);
    Result[5, 0].Value := MinMeanBox.Text;
    Result[5, 1].Name := 'Maximum item mean';
    // Result[5, 1].Value:= FormatFloat('0.00',MainForm.maxmeanbox.value);
    Result[5, 1].Value := MaxMeanBox.Text;
    Result[6, 0].Name := 'Minimum item correlation';
    Result[6, 0].Value := formatfloat('0.00', FMinR);
    Result[6, 1].Name := 'Maximum item correlation';
    Result[6, 1].Value := formatfloat('0.00', FMaxR);
    Result[7, 0].Name := 'ITEMAN 3.0 Header';
    // Result[7, 0].Value:= BoolToStr(MainForm.itapinc.checked, 'Yes', 'No');
    Result[7, 0].Value := BoolToStr(cbDataMatrixFile.IsChecked, True);
    Result[7, 1].Name := 'Exclude omits from option statistics';
    Result[7, 1].Value := BoolToStr(OmitAsinc.IsChecked, True);
    Result[8, 0].Name := 'Number of ID columns';
    Result[8, 0].Value := IntToStr(ExamineeIDColumns);

    Result[8, 1].Name := 'ID begins in column';
    if ExamineeIDColumns = 0 then
      Result[8, 1].Value := IDColumnsBegin.Text
    else
      Result[8, 1].Value := IfThen(cbDataMatrixFile.IsChecked, '1', IDColumnsBegin.Text);

    Result[9, 0].Name := 'Responses begin in column';
    Result[9, 0].Value := IfThen(cbDataMatrixFile.IsChecked, IntToStr(ExamineeIDColumns + 1), itemColumnBox.Text);
    Result[9, 1].Name := 'Omit character';
    Result[9, 1].Value := IfThen(cbDataMatrixFile.IsChecked, FOmitChar, omitcharbox.Text[1]);
    Result[10, 0].Name := 'Not Admin character';
    Result[10, 0].Value := IfThen(cbDataMatrixFile.IsChecked, FNAChar, nacharbox.Text[1]);
    Result[10, 1].Name := 'Produce quantile tables';
    Result[10, 1].Value := IfThen(tablebox.IsChecked, 'Yes', 'No');
    Result[11, 0].Name := 'Correct for spuriousness';
    Result[11, 0].Value := IfThen(spurbox.IsChecked, 'Yes', 'No');
    Result[11, 1].Name := 'Produce quantile plots';
    Result[11, 1].Value := IfThen(plotsbox.IsChecked, 'Yes', 'No');
    Result[12, 0].Name := 'Save data matrix';
    Result[12, 0].Value := IfThen(ScoredMatrixBox.IsChecked, 'Yes', 'No');
    Result[12, 1].Name := 'Include omit codes in matrix';
    Result[12, 1].Value := IfThen(ScoredMatrixBox.IsChecked, IfThen(InclOmit.IsChecked, 'Yes', 'No'), 'N/A');
    Result[13, 0].Name := 'Include Not Admin codes in matrix';
    Result[13, 0].Value := IfThen(ScoredMatrixBox.IsChecked, IfThen(InclNR.IsChecked, 'Yes', 'No'), 'N/A');

    Result[13, 1].Name := 'Include scaled scores for';
    if ScaledCheckBox.IsChecked and ScaledDomain.IsChecked then
      // if (cbscaledcheckbox.checked) and (MainForm.scaleddomain.checked) then
      Result[13, 1].Value := 'Total/Domain'
    else if ScaledDomain.IsChecked then
      Result[13, 1].Value := 'Domain Scores'
    else if ScaledCheckBox.IsChecked then
      Result[13, 1].Value := 'Total Scores'
    else
      Result[13, 1].Value := 'N/A';

    Result[14, 0].Name := 'Scaling function';
    Result[14, 0].Value := IfThen(linearscale.IsChecked or standscale.IsChecked,
      IfThen(linearscale.IsChecked, 'Linear', 'Standardized'), 'N/A');

    if linearscale.IsChecked then
    begin
      Result[14, 1].Name := 'Scaled score intercept';
      Result[14, 1].Value := interceptbox.Text;
      Result[15, 0].Name := 'Scaled score slope';
      Result[15, 0].Value := slopebox.Text;
    end
    else if standscale.IsChecked then
    begin
      Result[14, 1].Name := 'Scaled score new mean';
      Result[14, 1].Value := newmeanbox.Text;
      Result[15, 0].Name := 'Scaled score new SD';
      Result[15, 0].Value := newsdbox.Text;
    end
    else
    begin
      Result[14, 1].Name := 'Scaled score setting 1';
      Result[14, 1].Value := 'N/A';
      Result[15, 0].Name := 'Scaled score new SD';
      Result[15, 0].Value := 'N/A';
    end;

    Result[15, 1].Name := 'Dichotomous Classification';
    Result[15, 1].Value := IfThen(ScoredClass.IsChecked or ScaledClass.IsChecked,
      IfThen(ScoredClass.IsChecked, 'Total Score', 'Scaled Score'), 'N/A');
    Result[16, 0].Name := 'Classify based on';
    Result[16, 0].Value := IfThen(ScoredClass.IsChecked or ScaledClass.IsChecked,
      IfThen(ScoredClass.IsChecked, 'Total Score', 'Scaled Score'), 'N/A');
    Result[16, 1].Name := 'Cutpoint';
    Result[16, 1].Value := IfThen(classcheckbox.IsChecked, classcutbox.Text, 'N/A');
    Result[17, 0].Name := 'Low group label';
    Result[17, 0].Value := LowLabel.Text;
    Result[17, 1].Name := 'High group label';
    Result[17, 1].Value := HighLabel.Text;
    Result[18, 0].Name := 'Data is delimited by';
    Result[18, 0].Value := IfThen(commabox.IsChecked or tabbox.IsChecked,
      IfThen(commabox.IsChecked, 'Comma', 'Tab'), 'N/A');
    Result[18, 1].Name := 'Test for DIF';
    Result[18, 1].Value := IfThen(difcheckbox.IsChecked, 'Yes', 'No');
    Result[19, 0].Name := 'Group status is in column';
    Result[19, 0].Value := IfThen(difcheckbox.IsChecked, GroupColumn.Text, 'N/A');
    Result[19, 1].Name := 'Ability levels for DIF';
    Result[19, 1].Value := IfThen(difcheckbox.IsChecked, NumDifGroups.Text, 'N/A');
    Result[20, 0].Name := 'Group 1 code';
    Result[20, 0].Value := IfThen(difcheckbox.IsChecked, group1code.Text, 'N/A');
    Result[20, 1].Name := 'Group 2 code';
    Result[20, 1].Value := IfThen(difcheckbox.IsChecked, group2code.Text, 'N/A');
    Result[21, 0].Name := 'Group 1 label';
    Result[21, 0].Value := IfThen(difcheckbox.IsChecked, Group1label.Text, 'N/A');
    Result[21, 1].Name := 'Group 2 label';
    Result[21, 1].Value := IfThen(difcheckbox.IsChecked, Group2label.Text, 'N/A');
  end;
end;

function TItemanCalculation.IsNumeric(FCh: Char): Boolean;
var
  i: Integer;
  foundint: Boolean;
begin
  foundint := False;
  for i := 0 to 9 do
    if FCh = chr(48 + i) then
      foundint := True;
  if (FCh = chr(43)) or (FCh = chr(45)) then
    foundint := True; // don't error trap is control file is +/- poly
  Result := foundint;
END;

{ --------------------------Determine if character is alphabetic------------------ }
function TItemanCalculation.IsAlphaBetic(FCh: Char): Boolean;
var
  i: Integer;
  foundalp: Boolean;
begin
  foundalp := False;
  for i := 0 to 14 do
    if FCh = chr(65 + i) then
      foundalp := True;
  if (FCh = chr(43)) or (FCh = chr(45)) then
    foundalp := True; // don't error trap is control file is +/- poly
  Result := foundalp;
END;

{ This is needed because the response number is used as an index }
function TItemanCalculation.ConvertResponseToInteger(FCh: Char; iteminfoarray: TChar2Array;
  FItem, nalt: Integer): Integer;
var
  g: Integer;
begin
  Result := succ(nalt);

  if iteminfoarray[FItem, 3] = 'P' then
    for g := 0 to pred(nalt) do
      if (FCh = chr(48 + g)) then
        if iteminfoarray[FItem, 1] = '-' then
          Result := nalt - g - 1
        else
          Result := g; // chr(48) equals 0

  if iteminfoarray[FItem, 3] <> 'P' then
  begin
    for g := 1 to nalt do
      if (FCh = chr(48 + g)) or (FCh = chr(64 + g)) then // 65 is character 'A' in ASCII hence 64+g
        if iteminfoarray[FItem, 1] = '-' then
          Result := succ(nalt) - g
        else
          Result := g;

    if ((FCh = chr(48)) or (FCh = chr(26))) and ((FCh <> FOmitChar) and (FCh <> FNAChar)) then
      FAllOk := False;
  end; // if begin0 = '1' // change made Mar.2013 for error trapping part (allok := false changed to activate only when both na and omit <> FCh)

  if (Result = succ(nalt)) and (FOmitChar <> chr(48)) and (FNAChar <> chr(48)) then
    FRespOK := False; // chr(48) is 0
end;

function TItemanCalculation.RespToInt(FCh: Char; iteminfoarray: TChar2Array; FItem: Integer;
  itemscalearray: TIntegerArray): Integer;
var
  aj: Integer;
begin
  if iteminfoarray[FItem, 3] = 'P' then
    aj := 0
  else
    aj := 1;

  Result := 17 - aj; // starting value

  if FCh = '0' then
    Result := 1 - aj
  else if (FCh = '1') or (FCh = 'A') then
    Result := 2 - aj
  else if (FCh = '2') or (FCh = 'B') then
    Result := 3 - aj
  else if (FCh = '3') or (FCh = 'C') then
    Result := 4 - aj
  else if (FCh = '4') or (FCh = 'D') then
    Result := 5 - aj
  else if (FCh = '5') or (FCh = 'E') then
    Result := 6 - aj
  else if (FCh = '6') or (FCh = 'F') then
    Result := 7 - aj
  else if (FCh = '7') or (FCh = 'G') then
    Result := 8 - aj
  else if (FCh = '8') or (FCh = 'H') then
    Result := 9 - aj
  else if (FCh = '9') or (FCh = 'I') then
    Result := 10 - aj
  else if FCh = 'J' then
    Result := 11 - aj
  else if FCh = 'K' then
    Result := 12 - aj
  else if FCh = 'L' then
    Result := 13 - aj
  else if FCh = 'M' then
    Result := 14 - aj
  else;
  if FCh = 'N' then
    Result := 15 - aj
  else if FCh = 'O' then
    Result := 16 - aj;
  if UpCase(FCh) = UpCase(FOmitChar) then
    Result := succ(itemscalearray[FItem, 1])
  else if UpCase(FCh) = UpCase(FNAChar) then
    Result := itemscalearray[FItem, 1] + 2;
end;

procedure TItemanCalculation.mantelH(item: Integer; responsearray, iteminfoarray: TChar2Array;
  difrankarray: TIntegerArray; difarray: TRealArray);
var
  i, j, si, row, numcut: Integer;
  alphas: array [1 .. 7] of Real;
  contingency: array [1 .. 2, 1 .. 2] of Real;
  rk, sk, r, s, pk, qk, temp1, temp2, temp3: Real;
  // terms for the SE of M-H statistics
begin
  difarray[item, 1] := 0;
  numcut := StrToIntDef(frmMain.NumDifGroups.Text, 0);

  for si := 1 to 5 + numcut do
    difarray[item, si] := 0;

  temp1 := 0;
  temp2 := 0;
  temp3 := 0;
  r := 0;
  s := 0;

  for si := 1 to numcut do
  begin
    for i := 1 to 2 do
      for j := 1 to 2 do
        contingency[i, j] := 0;

    for row := 1 to FNumberOfExaminees do
      if difrankarray[row, 2] = si then
      begin
        if (responsearray[row, item] = iteminfoarray[item, 1]) and (responsearray[row, 0] = '1') then
          contingency[1, 1] := contingency[1, 1] + 1;
        if (responsearray[row, item] <> iteminfoarray[item, 1]) and (responsearray[row, 0] = '1') then
          contingency[1, 2] := contingency[1, 2] + 1;
        if (responsearray[row, item] = iteminfoarray[item, 1]) and (responsearray[row, 0] = '2') then
          contingency[2, 1] := contingency[2, 1] + 1;
        if (responsearray[row, item] <> iteminfoarray[item, 1]) and (responsearray[row, 0] = '2') then
          contingency[2, 2] := contingency[2, 2] + 1;
      end;

    for i := 1 to 2 do
      for j := 1 to 2 do
        if contingency[i, j] = 0 then
          contingency[i, j] := 0.5;

    alphas[si] := (contingency[1, 1] * contingency[2, 2]) / (contingency[1, 2] * contingency[2, 1]);
    rk := (contingency[1, 1] * contingency[2, 2]) / FNumberForDif;
    sk := (contingency[1, 2] * contingency[2, 1]) / FNumberForDif;
    pk := (contingency[1, 1] + contingency[2, 2]) / FNumberForDif;
    qk := 1 - pk;
    r := r + rk;
    s := s + sk;
    temp1 := temp1 + pk * rk;
    temp2 := temp2 + pk * sk + qk * rk;
    temp3 := temp3 + qk * sk;

    difarray[item, 1] := difarray[item, 1] + alphas[si] * ((contingency[1, 2] * contingency[2, 1]) / FNumberForDif);
    difarray[item, 2] := difarray[item, 2] + (contingency[1, 2] * contingency[2, 1]) / FNumberForDif;
  end; // end for numcut

  difarray[item, 1] := difarray[item, 1] / difarray[item, 2];
  // alpha M-H
  difarray[item, 2] := -2.35 * ln(difarray[item, 1]);
  difarray[item, 3] := sqrt(0.5 * ((temp1 / (r * r)) + (temp2 / (r * s)) + (temp3 / (s * s))));
  // Std error of log(alpha M-H)
  difarray[item, 4] := -1 * ln(difarray[item, 1]) / difarray[item, 3];
  // multiply by -1 so - = ref higher, + = focal higher
  difarray[item, 5] := cdfn(difarray[item, 4]);
  // FP-value for DIF FZ-test

  for si := 1 to numcut do
    difarray[item, si + 5] := alphas[si];
  // sub-FGroup alphas, useful for identifying crossing vs uniform DIF
end;

function TItemanCalculation.cdfn(x: single): single;
// Cdfn integrates the normal density function to the deviate x.
var
  y: single;
begin
  y := x * 0.707107;
  Result := 0.5 * (erfn(y) + 1.0)
END;

function TItemanCalculation.erfn(x: single): single;
const
  a1 = 0.254830;
  a2 = -0.284497;
  a3 = 1.421414;
  a4 = -1.453152;
  a5 = 1.061405;
  FP = 0.327591;

var
  at: single;
  eat: single;
  es: single;
  t: single;
  y: single;
  y2: single;
begin
  if (x < 0.0) then
    es := -1.0
  else
    es := 1.0;

  y := abs(x);

  if (y < 0.000001) then
    Result := 0.0
  else if (y > 6.0) then
    Result := es
  else
  begin
    y2 := y * y;
    t := 1.0 / (1.0 + FP * y);
    at := ((a1 + (a2 + (a3 + (a4 + a5 * t) * t) * t) * t) * t);
    eat := at / exp(y2);
    Result := (1.0 - eat) * es;
  end;
end;

procedure TItemanCalculation.DIFgroups(difrankarray, scorearray: TIntegerArray;
  iteminfoarray, responsearray: TChar2Array);
var
  i, j, row, numscoreshigher, numcut: Integer;
  cutpts: Array [1 .. 6] of Integer;
begin
  FNumberForDif := 0;
  // Rank all scores - first column is an indexing table
  // Cycle through people and count how many scores were higher

  for row := 1 to FNumberOfExaminees do
    if responsearray[row, 0] <> '0' then
    begin
      numscoreshigher := 0;
      inc(FNumberForDif);

      for i := 1 to FNumberOfExaminees do
        // don't compare that examinee to themself
        if (responsearray[i, 0] <> '0') and (row <> i) and (scorearray[row, 1] < scorearray[i, 1]) then
          numscoreshigher := numscoreshigher + 1;

      difrankarray[row, 1] := numscoreshigher + 1;
    end;

  numcut := StrToIntDef(frmMain.NumDifGroups.Text, 0);

  // Determine Cut Points
  for j := 1 to pred(numcut) do
    cutpts[j] := round((j * FNumberForDif) / numcut);

  // Assign Groups
  for row := 1 to FNumberOfExaminees do
    if responsearray[row, 0] <> '0' then
    begin
      difrankarray[row, 2] := numcut;

      for j := 1 to pred(numcut) do
        if difrankarray[row, 1] > cutpts[j] then
          difrankarray[row, 2] := numcut - j;
    end;
end;

procedure TItemanCalculation.ScoreGroups(RankingArray, ReduceArray, scorearray: TIntegerArray;
  iteminfoarray, responsearray: TChar2Array);
var
  i, j, row, column: Integer;
  cutpts: Array [1 .. 6] of Integer;
  lNumberScoresHigher: Integer;
begin
  { Rank all scores - first column is an indexing table }
  { Cycle through people and count how many scores were higher }

  for row := 1 to FNumberOfExaminees do
  begin
    lNumberScoresHigher := 0;

    for i := 1 to FNumberOfExaminees do { don't compare that examinee to themself }
      if (row <> i) and (scorearray[row, 1] < scorearray[i, 1]) then
        inc(lNumberScoresHigher);

    RankingArray[row, 1] := succ(lNumberScoresHigher);
  end;

  { Determine Cut Points }
  for j := 1 to pred(FNumCut) do
    cutpts[j] := round((j * FNumberOfExaminees) / FNumCut);

  { Assign Groups }
  for row := 1 to FNumberOfExaminees do
  begin
    RankingArray[row, 2] := FNumCut;

    for j := 1 to pred(FNumCut) do
      if RankingArray[row, 1] > cutpts[j] then
        RankingArray[row, 2] := FNumCut - j;
  end;

  { Find number in each FGroup }
  for j := 1 to 7 do
    FGroupCounts[j] := 0;

  for row := 1 to FNumberOfExaminees do
    for j := 1 to 7 do
      if RankingArray[row, 2] = j then
        FGroupCounts[j] := FGroupCounts[j] + 1;

  for row := 1 to FNumberOfExaminees do
    for column := 1 to FNumberOfItems do
      if iteminfoarray[column, 2] <> 'N' then
      begin
        if responsearray[row, column] = FNAChar then
          ReduceArray[column, RankingArray[row, 2]] := ReduceArray[column, RankingArray[row, 2]] + 1;

        if responsearray[row, column] = FOmitChar then
          ReduceArray[column, RankingArray[row, 2]] := ReduceArray[column, RankingArray[row, 2]] + 1;
      end;
end;

procedure TItemanCalculation.SumStats(iteminfoarray: TChar2Array; scorearray: TIntegerArray);
var
  row, column: Integer;
  rescaled, temp1, temp2, temp3, scoresum1, scoresum2, scoresum3, scoress1, scoress2, scoress3: Real;
begin
  // _______________   Score Statistics  ___________________}
  for row := 1 to 4 do
    for column := 1 to 8 do
      FTestStatsArray[row, column] := 0;

  { Numbers of each type of item }
  FTestStatsArray[2, 1] := FPreTest + FScoredItems;
  FTestStatsArray[1, 1] := 0;
  FTestStatsArray[3, 1] := 0;

  for row := 1 to FNumberOfItems do
    if iteminfoarray[row, 2] <> 'N' then
    begin
      If iteminfoarray[row, 2] = 'Y' then
        FTestStatsArray[1, 1] := FTestStatsArray[1, 1] + 1;

      If iteminfoarray[row, 2] = 'P' then
        FTestStatsArray[3, 1] := FTestStatsArray[3, 1] + 1;
    end;

  // ________Calculate summary stats_____________}
  scoresum1 := 0;
  scoresum2 := 0;
  scoresum3 := 0;

  for row := 1 to FNumberOfExaminees do
  begin
    scoresum1 := scoresum1 + scorearray[row, 4];
    scoresum2 := scoresum2 + scorearray[row, 1];
    scoresum3 := scoresum3 + scorearray[row, 5];
  end;

  FTestStatsArray[2, 2] := scoresum1 / FNumberOfExaminees;
  FTestStatsArray[1, 2] := scoresum2 / FNumberOfExaminees;
  FTestStatsArray[3, 2] := scoresum3 / FNumberOfExaminees;

  scoress1 := 0;
  scoress2 := 0;
  scoress3 := 0;

  for row := 1 to FNumberOfExaminees do
  begin
    scoress1 := scoress1 + (FTestStatsArray[2, 2] - scorearray[row, 4]) * (FTestStatsArray[2, 2] - scorearray[row, 4]);
    scoress2 := scoress2 + (FTestStatsArray[1, 2] - scorearray[row, 1]) * (FTestStatsArray[1, 2] - scorearray[row, 1]);
    scoress3 := scoress3 + (FTestStatsArray[3, 2] - scorearray[row, 5]) * (FTestStatsArray[3, 2] - scorearray[row, 5]);
  end;

  if FNumberOfExaminees > 1 then
    FTestStatsArray[2, 3] := sqrt(scoress1 / (FNumberOfExaminees - 1));
  if FNumberOfExaminees > 1 then
    FTestStatsArray[1, 3] := sqrt(scoress2 / (FNumberOfExaminees - 1));
  if FNumberOfExaminees > 1 then
    FTestStatsArray[3, 3] := sqrt(scoress3 / (FNumberOfExaminees - 1));

  // --- Min and Max for all items, scored items, and FPreTest items---}
  temp1 := 100000;
  temp2 := 100000;
  temp3 := 100000;

  for row := 1 to FNumberOfExaminees do
  begin
    If scorearray[row, 4] < temp1 then
      temp1 := scorearray[row, 4];
    if scorearray[row, 1] < temp2 then
      temp2 := scorearray[row, 1];
    if scorearray[row, 5] < temp3 then
      temp3 := scorearray[row, 5];
  end;

  FTestStatsArray[2, 4] := temp1;
  FTestStatsArray[1, 4] := temp2;
  FTestStatsArray[3, 4] := temp3;

  temp1 := 0;
  temp2 := 0;
  temp3 := 0;

  for row := 1 to FNumberOfExaminees do
  begin
    If scorearray[row, 4] > temp1 then
      temp1 := scorearray[row, 4];
    if scorearray[row, 1] > temp2 then
      temp2 := scorearray[row, 1];
    if scorearray[row, 5] > temp3 then
      temp3 := scorearray[row, 5];
  end;

  FTestStatsArray[2, 5] := temp1;
  FTestStatsArray[1, 5] := temp2;
  FTestStatsArray[3, 5] := temp3;

  // _________Calculate summary stats for scaled scores__________}
  if frmMain.ScaledCheckBox.IsChecked then
  begin
    scoresum1 := 0;
    rescaled := 0;

    for row := 1 to FNumberOfExaminees do
    begin
      if frmMain.standscale.IsChecked then
      begin
        rescaled := scorearray[row, 1] * (StrToFloatDef(frmMain.newsdbox.Text, 0) / FTestStatsArray[1, 3]);
        rescaled := rescaled + (StrToFloatDef(frmMain.newmeanbox.Text, 0) - (StrToFloatDef(frmMain.newsdbox.Text,
          0) / FTestStatsArray[1, 3]) * FTestStatsArray[1, 2]);
      end;

      if frmMain.linearscale.IsChecked then
        rescaled := (scorearray[row, 1] * StrToFloatDef(frmMain.slopebox.Text, 0)) +
          StrToFloatDef(frmMain.interceptbox.Text, 0);

      FScaledScoreArray[row] := rescaled;
      scoresum1 := scoresum1 + FScaledScoreArray[row];
    end;

    FTestStatsArray[4, 2] := scoresum1 / FNumberOfExaminees;
    FScoreSS := 0;

    for row := 1 to FNumberOfExaminees do
    begin
      FScoreSS := FScoreSS + (FTestStatsArray[4, 2] - FScaledScoreArray[row]) *
        (FTestStatsArray[4, 2] - FScaledScoreArray[row]);
    end;

    FTestStatsArray[4, 3] := sqrt(FScoreSS / pred(FNumberOfExaminees));

    // --- Min and Max for scaled scores---}
    temp1 := 10000;

    for row := 1 to FNumberOfExaminees do
      if FScaledScoreArray[row] < temp1 then
        temp1 := FScaledScoreArray[row];

    FTestStatsArray[4, 4] := temp1;
    temp1 := 0;

    for row := 1 to FNumberOfExaminees do
      if FScaledScoreArray[row] > temp1 then
        temp1 := FScaledScoreArray[row];

    FTestStatsArray[4, 5] := temp1;
  end; { ScaledCheckBox.IsChecked }

  { _________Calculate summary stats for EXTERNAL scores__________ }
  If frmMain.cbExternalScoreFile.IsChecked then
  begin
    scoresum1 := 0;

    for row := 1 to FNumberOfExaminees do
      scoresum1 := scoresum1 + FExternalScoreArray[row];

    FTestStatsArray[5, 2] := scoresum1 / FNumberOfExaminees;
    FScoreSS := 0;

    for row := 1 to FNumberOfExaminees do
      FScoreSS := FScoreSS + (FTestStatsArray[5, 2] - FExternalScoreArray[row]) *
        (FTestStatsArray[5, 2] - FExternalScoreArray[row]);

    FTestStatsArray[5, 3] := sqrt(FScoreSS / (FNumberOfExaminees - 1));

    { --- Min and Max for external scores--- }
    temp1 := 10000;

    for row := 1 to FNumberOfExaminees do
      if FExternalScoreArray[row] < temp1 then
        temp1 := FExternalScoreArray[row];

    FTestStatsArray[5, 4] := temp1;
    temp1 := 0;

    for row := 1 to FNumberOfExaminees do
      if FExternalScoreArray[row] > temp1 then
        temp1 := FExternalScoreArray[row];

    FTestStatsArray[5, 5] := temp1;
  end; { ScaledCheckBox.IsChecked }
end;

procedure TItemanCalculation.OptionStats(iteminfoarray, responsearray: TChar2Array;
  itemscalearray, nkeyarray, scorearray: TIntegerArray);
var
  row, column, i, j, k, itemscore: Integer;
  correct: Boolean;
  ssbet, sstot, exts, temp0, temp1, temp2, temp3, temp4, rawsum, TMean, TSD: Real;
  multiN, multisd, multimean, multisum, multibis: Real;
  dommean: array [1 .. 16] of Real;
  domcount: array [1 .. 16] of Integer;
begin
  if not frmMain.OmitAsinc.IsChecked then
    FOmch := chr(7)
  else
    FOmch := FOmitChar; // if omits are scored replacing FOmitChar

  { ---Set Options Freqs to 0--- }
  for row := 1 to FNumberOfItems do { row = item, column = option }
    for column := 1 to 17 do
    begin
      FOptionCountsArray[row, column] := 0;
      FOptionRpbisArray[row, column] := 0;
      FOptionRbisArray[row, column] := 0;
      FOptionMeansArray[row, column] := 0;
      FOptionSDArray[row, column] := 0;
    end;

  FMinItem1 := 999;
  FMaxItem1 := -1;
  FMinRPBis := 5;
  FMinCorr := 5;
  maxcategory := 0;

  { ---Count Option Freqs--- }
  for column := 1 to FNumberOfItems do
    if iteminfoarray[column, 2] <> 'N' then { Column = item }
    begin
      temp0 := 0;
      temp1 := 0;
      temp2 := 0;
      frmMain.ProgressBar.Value := frmMain.ProgressBar.Value + 1;
      Application.ProcessMessages;

      for row := 1 to FNumberOfExaminees do { row = person }
      begin
        if responsearray[row, column] <> FOmitChar then
        begin
          if responsearray[row, column] <> FNAChar then
          begin
            FCh := responsearray[row, column];

            if row = 1 then
              for i := 1 to 10 do
                if FCh = chr(i + 47) then
                  FRespAsInt := True; // start at ascii 0 go to ascii 9

            if row = 1 then
              for i := 1 to 26 do
                if FCh = chr(i + 64) then
                  FRespAsInt := False; // start at ascii A go to ascii FZ

            FResponseAsInt := ConvertResponseToInteger(FCh, iteminfoarray, column, itemscalearray[column, 1]);

            if iteminfoarray[column, 3] = 'P' then
              FWts := 1
            else
              FWts := 0;

            FOptionCountsArray[column, FResponseAsInt + FWts] := FOptionCountsArray[column, FResponseAsInt + FWts] + 1;

            if (iteminfoarray[column, 3] = 'M') and (iteminfoarray[column, 1] = UpCase(FCh)) then
              temp0 := temp0 + 1;

            if (nkeyarray[column, 1] > 1) then
              for k := 1 to nkeyarray[column, 1] - 1 do
                if UpCase(iteminfoarray[column, k + 3]) = UpCase(FCh) then
                  temp0 := temp0 + 1;

            if iteminfoarray[column, 3] <> 'M' then
              temp0 := temp0 + FResponseAsInt;
          end; // if <> not reached
        end; // if <> omit

        if responsearray[row, column] = FOmitChar then
          FOptionCountsArray[column, 16] := FOptionCountsArray[column, 16] + 1;
        if responsearray[row, column] = FNAChar then
          FOptionCountsArray[column, 17] := FOptionCountsArray[column, 17] + 1;
      end;

      { ---Calculate option proportions--- }
      FItemN := 0;

      for row := 1 to FMaxItemOptions do
        FItemN := FItemN + FOptionCountsArray[column, row];

      if not frmMain.OmitAsinc.IsChecked and (iteminfoarray[column, 9] <> 'P') then
        FItemN := FItemN + FOptionCountsArray[column, 16];

      FItemStatsArray[column, 2] := temp0 / (FItemN + 1E-20);
      // store the item mean, independent of item type
      FItemStatsArray[column, 1] := FItemN;
      { record in overall item stat array }

      if iteminfoarray[column, 3] <> 'M' then
      begin
        if FItemStatsArray[column, 2] < FMinItem1 then
          FMinItem1 := FItemStatsArray[column, 2];
        if FItemStatsArray[column, 2] > FMaxItem1 then
          FMaxItem1 := FItemStatsArray[column, 2];
        if itemscalearray[column, 1] > maxcategory then
          maxcategory := itemscalearray[column, 1];
      end;

      if iteminfoarray[column, 3] = 'M' then
        for row := 1 to FNumberOfExaminees do
        begin
          if (responsearray[row, column] <> FNAChar) and (responsearray[row, column] <> FOmch) then
            if UpCase(iteminfoarray[column, 1]) = UpCase(responsearray[row, column]) then
              temp1 := temp1 + (1 - FItemStatsArray[column, 2]) * (1 - FItemStatsArray[column, 2])
            else
              temp1 := temp1 + (0 - FItemStatsArray[column, 2]) * (0 - FItemStatsArray[column, 2]);

          if (nkeyarray[column, 1] > 1) then
            for k := 1 to nkeyarray[column, 1] - 1 do
              if UpCase(iteminfoarray[column, k + 3]) = UpCase(FCh) then
                temp1 := temp1 + (1 - FItemStatsArray[column, 2]) * (1 - FItemStatsArray[column, 2]);
        end; // for row loop for multiple choice items

      if iteminfoarray[column, 3] = 'M' then
        FItemStatsArray[column, 6] := temp1 / (FItemN + 1E-20);

      for row := 1 to FMaxItemOptions do
        FOptionPropsArray[column, row] := FOptionCountsArray[column, row] / (FItemN + 0.000000000000001);

      if not frmMain.OmitAsinc.IsChecked then
        FOptionPropsArray[column, 16] := FOptionCountsArray[column, 16] / (FItemN + 1E-20);
    end; // FNumberOfItems loop

  { ---Calculate option rpbis--- }
  for i := 1 to FNumberOfItems do
    if iteminfoarray[i, 2] <> 'N' then
    begin
      frmMain.ProgressBar.Value := frmMain.ProgressBar.Value + 1;
      Application.ProcessMessages;

      if iteminfoarray[i, 3] = 'P' then
        FWts := 1
      else
        FWts := 0;

      for j := 1 to FMaxItemOptions do
      begin
        { Reset all sums to zero }
        FRpbisCalcArray[j, 1] := 0;
        FRpbisCalcArray[j, 2] := 0;
        FRpbisCalcArray[16, 1] := 0;
        FRpbisCalcArray[17, 1] := 0;

        { Calculate corrected score mean and SD for full Rpbis }
        FScoreSum := 0;
        FScoreSS := 0;
        FSpurN := 0;
        rawsum := 0;

        for row := 1 to FNumberOfExaminees do
        begin
          if (responsearray[row, i] <> FOmch) and (responsearray[row, i] <> FNAChar) then
          begin
            { First, correct score (in ScoreArray) for spuriousness }
            scorearray[row, 3] := scorearray[row, 1];
            FCh := responsearray[row, i];

            if frmMain.spurbox.IsChecked then { If correcting for spur - if not, scores are left as is }
            begin
              if (iteminfoarray[i, 3] = 'M') and (UpCase(FCh) = UpCase(iteminfoarray[i, 1]))
              then { If the response is the FKey }
                if iteminfoarray[i, 2] = 'Y' then { If item is scored }
                  scorearray[row, 3] := scorearray[row, 3] - 1;
              { Subtracts the option 1/0 indicator }

              if (nkeyarray[i, 1] > 1) and (iteminfoarray[i, 2] = 'Y') then
                for k := 1 to nkeyarray[i, 1] - 1 do
                  if UpCase(iteminfoarray[i, k + 3]) = UpCase(FCh) then
                    scorearray[row, 3] := scorearray[row, 3] - 1;

              if (iteminfoarray[i, 3] <> 'M') and (iteminfoarray[i, 2] = 'Y') then
                scorearray[row, 3] := scorearray[row, 3] - ConvertResponseToInteger(responsearray[row, i],
                  iteminfoarray, i, itemscalearray[i, 1]);
            end;

            { Add that corrected score to running sum for the total corrected mean }
            if not frmMain.cbExternalScoreFile.IsChecked then
              FScoreSum := FScoreSum + scorearray[row, 3]
            else
              FScoreSum := FScoreSum + FExternalScoreArray[row];

            FSpurN := FSpurN + 1;

            { Add to running sum for option means }
            FResponseAsInt := ConvertResponseToInteger(FCh, iteminfoarray, i, itemscalearray[i, 1]);

            If FResponseAsInt + FWts = j then
            begin
              if not frmMain.cbExternalScoreFile.IsChecked then
                FRpbisCalcArray[j, 1] := FRpbisCalcArray[j, 1] + scorearray[row, 3];
              if frmMain.cbExternalScoreFile.IsChecked then
                FRpbisCalcArray[j, 1] := FRpbisCalcArray[j, 1] + FExternalScoreArray[row];
              if not frmMain.cbExternalScoreFile.IsChecked then
                rawsum := rawsum + scorearray[row, 1];
              if frmMain.cbExternalScoreFile.IsChecked then
                rawsum := rawsum + FExternalScoreArray[row];
            end;
          end;

          if frmMain.cbExternalScoreFile.IsChecked then
            exts := FExternalScoreArray[row]
          else
            exts := scorearray[row, 1];
          if (iteminfoarray[i, 2] = 'Y') and (responsearray[row, i] = FOmitChar) then
            FRpbisCalcArray[16, 1] := FRpbisCalcArray[16, 1] + exts;
          if (iteminfoarray[i, 2] = 'P') and (responsearray[row, i] = FOmitChar) then
            FRpbisCalcArray[16, 1] := FRpbisCalcArray[16, 1] + exts;
          if (iteminfoarray[i, 2] = 'Y') and (responsearray[row, i] = FNAChar) then
            FRpbisCalcArray[17, 1] := FRpbisCalcArray[17, 1] + exts;
          if (iteminfoarray[i, 2] = 'P') and (responsearray[row, i] = FNAChar) then
            FRpbisCalcArray[17, 1] := FRpbisCalcArray[17, 1] + exts;
        end; // for row loop from 1 to FNumberOfExaminees

        FSpurScoreMean := FScoreSum / (FSpurN + 0.0000000000000001);

        { Calculate option means }
        FOptionMeansArray[i, j] := rawsum / (FOptionCountsArray[i, j] + 1E-20);
        FOptionMeansArray[i, 16] := (FRpbisCalcArray[16, 1] / (FOptionCountsArray[i, 16] + 0.0000000000001));
        FOptionMeansArray[i, 17] := (FRpbisCalcArray[17, 1] / (FOptionCountsArray[i, 17] + 0.0000000000001));

        multiN := 0;
        multisd := 0;
        multimean := 0;
        multisum := 0;
        exts := 0;

        for row := 1 to FNumberOfExaminees do
        begin
          if (responsearray[row, i] <> FOmch) and (responsearray[row, i] <> FNAChar) then
          begin
            multiN := multiN + 1;

            if frmMain.cbExternalScoreFile.IsChecked then
              exts := FExternalScoreArray[row]
            else
              exts := scorearray[row, 3];

            FScoreSS := FScoreSS + (FSpurScoreMean - exts) * (FSpurScoreMean - exts);
          end;

          multimean := multimean + scorearray[row, 3];
          FResponseAsInt := ConvertResponseToInteger(responsearray[row, i], iteminfoarray, i, itemscalearray[i, 1]);

          if FResponseAsInt + FWts = j then
            FOptionSDArray[i, j] := FOptionSDArray[i, j] + (exts - FOptionMeansArray[i, j]) *
              (exts - FOptionMeansArray[i, j]);

          if frmMain.cbExternalScoreFile.IsChecked then
            exts := FExternalScoreArray[row]
          else
            exts := scorearray[row, 1]; // now full score
          if (responsearray[row, i] = FOmitChar) and (j = 1) then
            FOptionSDArray[i, 16] := FOptionSDArray[i, 16] + (exts - FOptionMeansArray[i, 16]) *
              (exts - FOptionMeansArray[i, 16]);
          if (responsearray[row, i] = FNAChar) and (j = 1) then
            FOptionSDArray[i, 17] := FOptionSDArray[i, 17] + (exts - FOptionMeansArray[i, 17]) *
              (exts - FOptionMeansArray[i, 17]);
        end;

        // Multikeyed items???? ---------------------------------
        if nkeyarray[i, 1] > 1 then
        begin
          FP := FItemStatsArray[i, 2];
          multimean := multimean / (multiN + 1E-20);

          for row := 1 to FNumberOfExaminees do
            if (responsearray[row, i] <> FOmch) and (responsearray[row, i] <> FNAChar) then
            begin
              if frmMain.cbExternalScoreFile.IsChecked then
                exts := FExternalScoreArray[row]
              else
                exts := scorearray[row, 3];

              multisd := multisd + (exts - multimean) * (exts - multimean);
            end;

          multisd := sqrt(multisd / (multiN - 1));

          for row := 1 to FNumberOfExaminees do
            if (responsearray[row, i] <> FOmch) and (responsearray[row, i] <> FNAChar) then
            begin
              correct := False;
              FCh := responsearray[row, i];

              if frmMain.cbExternalScoreFile.IsChecked then
                exts := FExternalScoreArray[row]
              else
                exts := scorearray[row, 3];

              if UpCase(iteminfoarray[i, 1]) = UpCase(FCh) then
                multisum := multisum + ((exts - multimean) / (multisd + 1E-20)) *
                  ((1 - FP) / (sqrt(FP * (1 - FP)) + 1E-20))
              else
                for k := 1 to nkeyarray[i, 1] - 1 do
                  if UpCase(iteminfoarray[i, k + 3]) = UpCase(FCh) then
                    multisum := multisum + ((exts - multimean) / (multisd + 1E-20)) *
                      ((1 - FP) / (sqrt(FP * (1 - FP)) + 1E-20));

              if (UpCase(iteminfoarray[i, 1]) = UpCase(FCh)) then
                correct := True;

              for k := 1 to nkeyarray[i, 1] - 1 do
                if UpCase(iteminfoarray[i, k + 3]) = UpCase(FCh) then
                  correct := True;

              if not correct then
                multisum := multisum + ((exts - multimean) / (multisd + 1E-20)) *
                  ((0 - FP) / (sqrt(FP * (1 - FP)) + 1E-20));
            end;

          FItemStatsArray[i, 3] := multisum / (multiN - 1 + 1E-20);
          { Convert Rpbis to Rbis }
          FZ := -cdfni(FP);
          FOrdinate := 0.398942 * exp(-sqr(FZ) / 2);
          multibis := FItemStatsArray[i, 3] * (sqrt(FP * (1 - FP)) / FOrdinate);

          If multibis > 1.0 then
            multibis := 1.0; { cap }
          If multibis < -1.0 then
            multibis := -1.0; { cap }
          FItemStatsArray[i, 4] := multibis;
        end; // if nkeyarray > 1

        if FSpurN > 1 then
          FSpurScoreSD := sqrt(FScoreSS / (FSpurN - 1));

        if FOptionCountsArray[i, j] > 1 then
          FOptionSDArray[i, j] := power((FOptionSDArray[i, j] / (FOptionCountsArray[i, j] - 1)), 0.5);
        if (FOptionCountsArray[i, 16] > 1) and (j = 1) then
          FOptionSDArray[i, 16] := power((FOptionSDArray[i, 16] / (FOptionCountsArray[i, 16] - 1)), 0.5);
        if (FOptionCountsArray[i, 17] > 1) and (j = 1) then
          FOptionSDArray[i, 17] := power((FOptionSDArray[i, 17] / (FOptionCountsArray[i, 17] - 1)), 0.5);

        { Calculate rpbis for each option - full method }
        FP := FOptionPropsArray[i, j];
        FRpbisSum := 0;
        FNforRpbis := 0;

        if FP > 0 then
        begin
          // Original code before debugging 2017.06.07
          for row := 1 to FNumberOfExaminees do
            if (responsearray[row, i] <> FOmch) and (responsearray[row, i] <> FNAChar) then
            begin
              FNforRpbis := FNforRpbis + 1;

              if frmMain.cbExternalScoreFile.IsChecked then
                exts := FExternalScoreArray[row]
              else
                exts := scorearray[row, 3];

              if ConvertResponseToInteger(responsearray[row, i], iteminfoarray, i, itemscalearray[i, 1]) + FWts = j then
                FRpbisSum := FRpbisSum + ((exts - FSpurScoreMean) / (FSpurScoreSD + 0.000000000000000000001)) *
                  ((1 - FP) / (sqrt(FP * (1 - FP)) + 0.000000000000000000001))
              else
                FRpbisSum := FRpbisSum + ((exts - FSpurScoreMean) / (FSpurScoreSD + 0.000000000000000000001)) *
                  ((0 - FP) / (sqrt(FP * (1 - FP)) + 0.000000000000000000001));
            end;

          FOptionRpbisArray[i, j] := FRpbisSum / (FNforRpbis - 1 + 1E-20);

          if (FOptionRpbisArray[i, j] > 1) then
            FOptionRpbisArray[i, j] := 1;

          { Convert Rpbis to Rbis }
          FZ := -cdfni(FP);
          FOrdinate := 0.398942 * exp(-sqr(FZ) / 2);
          FOptionRbisArray[i, j] := FOptionRpbisArray[i, j] * (sqrt(FP * (1 - FP)) / FOrdinate);

          If FOptionRbisArray[i, j] > 1.0 then
            FOptionRbisArray[i, j] := 1.0; { cap }
          If FOptionRbisArray[i, j] < -1.0 then
            FOptionRbisArray[i, j] := -1.0; { cap }
        end; // if FP > 0
      end; { j = options }

      { -----------------------Calculate rpbis for omitted response------------------------- }
      if frmMain.OmitAsinc.IsChecked = False then
      begin
        for row := 1 to FNumberOfExaminees do
          if (responsearray[row, i] <> FOmch) and (responsearray[row, i] <> FNAChar) then
          begin
            if frmMain.cbExternalScoreFile.IsChecked then
              exts := FExternalScoreArray[row]
            else
              exts := scorearray[row, 3];

            FScoreSS := FScoreSS + (FOptionMeansArray[i, 16] - exts) * (FOptionMeansArray[i, 16] - exts);
          end;

        if FSpurN > 1 then
          FSpurScoreSD := sqrt(FScoreSS / (FSpurN - 1));

        FP := FOptionPropsArray[i, 16];
        FRpbisSum := 0;
        FNforRpbis := 0;

        if FP > 0 then
        begin
          for row := 1 to FNumberOfExaminees do
          begin
            If responsearray[row, i] <> FNAChar then
            begin
              FNforRpbis := FNforRpbis + 1;

              if frmMain.cbExternalScoreFile.IsChecked then
                exts := FExternalScoreArray[row]
              else
                exts := scorearray[row, 3];

              FRpbisSum := FRpbisSum + ((exts - FOptionMeansArray[i, 16]) / (FSpurScoreSD + 0.000000000000000000001)) *
                ((0 - FP) / (sqrt(FP * (1 - FP)) + 0.000000000000000000001));
            end;
          end;
          FOptionRpbisArray[i, 16] := FRpbisSum / (FNforRpbis - 1 + 1E-20);

          { Convert Rpbis to Rbis }
          FZ := -cdfni(FP);
          FOrdinate := 0.398942 * exp(-sqr(FZ) / 2);
          FOptionRbisArray[i, 16] := FOptionRpbisArray[i, 16] * (sqrt(FP * (1 - FP)) / FOrdinate);

          if FOptionRbisArray[i, 16] > 1.0 then
            FOptionRbisArray[i, 16] := 1.0; { cap }
          if FOptionRbisArray[i, 16] < -1.0 then
            FOptionRbisArray[i, 16] := -1.0; { cap }
        end; // if FP > 0
      end; // if omitasinc.IsChecked
    end; { i = items }

  for column := 1 to FNumberOfItems do
    if (iteminfoarray[column, 2] <> 'N') and (iteminfoarray[column, 3] <> 'M') then
    begin
      temp0 := 0;
      temp1 := 0;
      temp2 := 0;
      temp4 := 0;
      TMean := 0;
      TSD := 0;
      FItemN := 0;
      ssbet := 0;

      for row := 1 to 16 do
        dommean[row] := 0; // reset to 0
      for row := 1 to 16 do
        domcount[row] := 0; // reset to 0

      for row := 1 to FNumberOfExaminees do
        if (responsearray[row, column] <> FOmitChar) and (responsearray[row, column] <> FNAChar) then
        begin
          FItemN := FItemN + 1;
          FCh := responsearray[row, column];
          FResponseAsInt := ConvertResponseToInteger(FCh, iteminfoarray, column, itemscalearray[column, 1]);

          if (frmMain.spurbox.IsChecked = False) or (iteminfoarray[column, 2] = 'P') then
            FResponseAsInt := 0;

          TMean := (scorearray[row, 1] - FResponseAsInt) + TMean;
        end;

      TMean := TMean / (FItemN + 1E-20);

      for row := 1 to FNumberOfExaminees do
        if (responsearray[row, column] <> FOmitChar) and (responsearray[row, column] <> FNAChar) then
        begin
          FCh := responsearray[row, column];
          FResponseAsInt := ConvertResponseToInteger(FCh, iteminfoarray, column, itemscalearray[column, 1]);

          if (frmMain.spurbox.IsChecked = False) or (iteminfoarray[column, 2] = 'P') then
            FResponseAsInt := 0;

          TSD := TSD + ((scorearray[row, 1] - FResponseAsInt) - TMean) *
            ((scorearray[row, 1] - FResponseAsInt) - TMean);
        end;

      if FItemN > 1 then
        TSD := sqrt(TSD / (FItemN - 1));

      for row := 1 to FNumberOfExaminees do
      begin
        FCh := responsearray[row, column];
        FResponseAsInt := ConvertResponseToInteger(FCh, iteminfoarray, column, itemscalearray[column, 1]);

        if (responsearray[row, column] = FOmitChar) and (responsearray[row, column] = FNAChar) then
          FResponseAsInt := 0;

        if (responsearray[row, column] <> FOmitChar) and (responsearray[row, column] <> FNAChar) then
        begin
          if (frmMain.spurbox.IsChecked = False) or (iteminfoarray[column, 2] = 'P') then
            itemscore := 0
          else
            itemscore := FResponseAsInt;

          if frmMain.spurbox.IsChecked then
            temp4 := temp4 + (scorearray[row, 1] - itemscore - TMean) * (scorearray[row, 1] - itemscore - TMean);

          dommean[FResponseAsInt + FWts] := dommean[FResponseAsInt + FWts] + (scorearray[row, 1] - itemscore);
          domcount[FResponseAsInt + FWts] := domcount[FResponseAsInt + FWts] + 1;
          temp1 := temp1 + (FResponseAsInt - FItemStatsArray[column, 2]) *
            (FResponseAsInt - FItemStatsArray[column, 2]);
          temp2 := temp2 + ((FResponseAsInt - FItemStatsArray[column, 2]) * (scorearray[row, 1] - itemscore - TMean));
        end; // FNumberOfExaminees loop
      end; // if response is not omit FNAChar

      for row := 1 to itemscalearray[column, 1] do
        dommean[row] := dommean[row] / (domcount[row] + 1E-20);

      if frmMain.spurbox.IsChecked then
        sstot := temp4
      else
        sstot := TSD * TSD * (FItemN - 1); // for eta

      if FItemN > 1 then
        temp4 := temp4 / (FItemN - 1);
      if FItemN > 1 then
        FItemStatsArray[column, 6] := temp1 / (FItemN - 1);

      for row := 1 to itemscalearray[column, 1] do
        ssbet := ssbet + domcount[row] * (dommean[row] - (TMean - FItemStatsArray[column, 2])) *
          (dommean[row] - (TMean - FItemStatsArray[column, 2]));

      if frmMain.spurbox.IsChecked then
        temp3 := power(FItemStatsArray[column, 6], 0.5) * power(temp4, 0.5) * (FItemN - 1)
      else
        temp3 := power(FItemStatsArray[column, 6], 0.5) * TSD * (FItemN - 1);

      FItemStatsArray[column, 3] := temp2 / (temp3 - 1 + 1E-20);

      if FItemStatsArray[column, 3] < FMinCorr then
        FMinCorr := FItemStatsArray[column, 3];

      FItemStatsArray[column, 4] := sqrt(ssbet / (sstot + 1E-20));
    end; // if item is 'R' or polytomous/rating scale

  { --- Find FP and rpbis for keyed options --- }
  for i := 1 to FNumberOfItems do
    if (iteminfoarray[i, 2] <> 'N') then
    begin
      if (iteminfoarray[i, 3] = 'M') then
      begin
        FCh := iteminfoarray[i, 1];
        FKey := ConvertResponseToInteger(FCh, iteminfoarray, i, itemscalearray[i, 1]);

        if nkeyarray[i, 1] = 1 then
          FItemStatsArray[i, 3] := FOptionRpbisArray[i, FKey];

        if (iteminfoarray[i, 2] = 'Y') and (FItemStatsArray[i, 3] < FMinRPBis) then
          FMinRPBis := FItemStatsArray[i, 3];

        if nkeyarray[i, 1] = 1 then
          FItemStatsArray[i, 4] := FOptionRbisArray[i, FKey];
      end; // if item is multiple choice

      if (iteminfoarray[i, 3] <> 'M') and (itemscalearray[i, 1] = 2) then
        FItemStatsArray[i, 4] := FOptionRbisArray[i, 2];
    end;
end;

procedure TItemanCalculation.PrintFirstSummary(prec: string; iteminfoarray: TChar2Array; itemidarray: TStringArray;
  itemscalearray: TIntegerArray; difarray: TRealArray);
var
  s: string;
  tbMatrix: TTableMatrix;
  h, i, j, v, coladj, flagrow, row: Integer;
  checkkey: Boolean;
  PreStr: string;
  LBitmap: TBitmap;
begin

  { Write description of Table 2 to rtf outpout }
  // AddSummary(frmMain.imgSumm.Bitmap);
  AddSummary(FSummBitmap);
  s := 'Table 2 presents the summary statistics of the test ';

  if (FPreTest > 0) and (FNumDomains <> 1) then
    s := s + 'for all items, scored items only, pretest items only,';
  if (FPreTest > 0) and (FNumDomains = 1) then
    s := s + 'for all items, scored items only, and pretest items only.';
  if (FPreTest = 0) and (FNumDomains <> 1) then
    s := s + 'for all items, scored items only,';
  if (FPreTest = 0) and (FNumDomains = 1) then
    s := s + 'for the scored items.';
  if FNumDomains <> 1 then
    s := s + ' and for each domain (content area).';
  s := s + ' Definitions of these statistics are found in the Iteman manual.';
  AddText(PWideChar(s));

  with frmMain do
  begin
    { Determine number of columns in summary table }
    if (FNumDomains = 1) and (FPreTest = 0) then
      summarytablerows := 2
    else
      summarytablerows := 4;
    If ScaledCheckBox.IsChecked then
      summarytablerows := summarytablerows + 1;
    if (FNumDomains > 1) and (ScaledDomain.IsChecked) then
      summarytablerows := summarytablerows + FNumDomains;
    If cbExternalScoreFile.IsChecked then
      summarytablerows := summarytablerows + 1;
    if (FPreTest = 0) and (FNumDomains <> 1) then
      summarytablerows := summarytablerows - 1;
    if FNumDomains <> 1 then
      summarytablerows := summarytablerows + FNumDomains;

    { Write summary statistics to rtf output AS A RICHVIEW TABLE }
    if FIsRate and FIsMult then
      coladj := 1
    else
      coladj := 0;
    SetLength(tbMatrix, summarytablerows, 8 + coladj);
    // SummaryTable := TRVTableItemInfo.CreateEx(SummaryTableRows, , RichView1.rvdata);
    tbMatrix[0, 0] := 'Score';
    tbMatrix[0, 1] := 'Items';
    tbMatrix[0, 2] := 'Mean';
    tbMatrix[0, 3] := 'SD';
    tbMatrix[0, 4] := 'Min Score';
    tbMatrix[0, 5] := 'Max Score';
    if FIsMult = True then
      tbMatrix[0, 6] := 'Mean P';
    if FIsRate = True then
      tbMatrix[0, 6 + coladj] := 'Item Mean';
    if FIsRate = True then
      tbMatrix[0, 7 + coladj] := 'Mean R'
    else
      tbMatrix[0, 7 + coladj] := 'Mean Rpbis';

    if (FNumDomains <> 1) or (FPreTest > 0) then
    begin
      tbMatrix[1, 0] := 'All items';
      tbMatrix[1, 1] := formatfloat('###', FTestStatsArray[2, 1]);
      for j := 2 to 3 do
        tbMatrix[1, j] := formatfloat(prec, FTestStatsArray[2, j]);
      for j := 4 to 5 do
        tbMatrix[1, j] := formatfloat('0', FTestStatsArray[2, j]);
      if FIsMult = True then
        tbMatrix[1, 6] := formatfloat(prec, FTestStatsArray[2, 7]);
      if FIsRate = True then
        tbMatrix[1, 6 + coladj] := formatfloat(prec, FTestStatsArray[2, 16]);
      tbMatrix[1, 7 + coladj] := formatfloat(prec, FTestStatsArray[2, 8]);
    end;

    if (FNumDomains <> 1) or (FPreTest > 0) then
      tempint := 2
    else
      tempint := 1;
    tbMatrix[tempint, 0] := 'Scored Items';
    tbMatrix[tempint, 1] := formatfloat('###', FTestStatsArray[1, 1]);
    for j := 2 to 3 do
      tbMatrix[tempint, j] := formatfloat(prec, FTestStatsArray[1, j]);
    for j := 4 to 5 do
      tbMatrix[tempint, j] := formatfloat('0', FTestStatsArray[1, j]);
    if FIsMult = True then
      tbMatrix[tempint, 6] := formatfloat(prec, FTestStatsArray[1, 7]);
    if FIsRate = True then
      tbMatrix[tempint, 6 + coladj] := formatfloat(prec, FTestStatsArray[1, 16]);
    tbMatrix[tempint, 7 + coladj] := formatfloat(prec, FTestStatsArray[1, 8]);

    if FPreTest > 0 then
    begin
      tbMatrix[3, 0] := 'Pretest items';
      tbMatrix[3, 1] := formatfloat('###', FTestStatsArray[3, 1]);
      for j := 2 to 3 do
        tbMatrix[3, j] := formatfloat(prec, FTestStatsArray[3, j]);
      for j := 4 to 5 do
        tbMatrix[3, j] := formatfloat('0', FTestStatsArray[3, j]);
      if FIsMult = True then
        tbMatrix[3, 6] := formatfloat(prec, FTestStatsArray[3, 7]);
      if FIsRate = True then
        tbMatrix[3, 6 + coladj] := formatfloat(prec, FTestStatsArray[3, 16]);
      tbMatrix[3, 7 + coladj] := formatfloat(prec, FTestStatsArray[3, 8]);
    end;

    tempint := 4;
    if FPreTest = 0 then
      tempint := tempint - 2;
    If cbExternalScoreFile.IsChecked then
    begin
      tbMatrix[tempint, 0] := 'External score';
      tbMatrix[tempint, 1] := '-';
      for j := 2 to 5 do
        tbMatrix[tempint, j] := formatfloat(prec, FTestStatsArray[5, j]);
      for j := 6 to 7 + coladj do
        tbMatrix[tempint, j] := '-';
    end;

    If cbExternalScoreFile.IsChecked then
    begin
      if FPreTest > 0 then
        Start := 5
      else
        Start := 4;
      if FPreTest > 0 then
        Finish := FNumDomains + 4
      else
        Finish := FNumDomains + 3;
    end
    else
    begin
      if FPreTest > 0 then
        Start := 4
      else
        Start := 3;
      if FPreTest > 0 then
        Finish := FNumDomains + 3
      else
        Finish := FNumDomains + 2;
    end;

    if FNumDomains <> 1 then
    begin
      Domain := 0;
      For row := Start to Finish do
      begin
        Domain := Domain + 1;
        tbMatrix[row, 0] := FStoreDomain[Domain];
        tbMatrix[row, 1] := formatfloat('###', FDomainStatsArray[Domain, 1]);
        for j := 2 to 3 do
          tbMatrix[row, j] := formatfloat(prec, FDomainStatsArray[Domain, j]);
        for j := 4 to 5 do
          tbMatrix[row, j] := formatfloat('0', FDomainStatsArray[Domain, j]);
        if FIsMult then
          tbMatrix[row, 6] := formatfloat(prec, FDomainStatsArray[Domain, 7]);
        if FIsRate then
          tbMatrix[row, 6 + coladj] := formatfloat(prec, FDomainStatsArray[Domain, 16]);
        tbMatrix[row, 7 + coladj] := formatfloat(prec, FDomainStatsArray[Domain, 8]);
      end;
    end;
    if FNumDomains = 1 then
      Finish := summarytablerows - 1;
    if (FNumDomains <> 1) and (ScaledDomain.IsChecked) then
      Finish := summarytablerows - 2 - FNumDomains;
    if (FNumDomains <> 1) and (ScaledDomain.IsChecked = False) then
      Finish := summarytablerows - 1;
    If ScaledCheckBox.IsChecked then
    begin
      if ScaledDomain.IsChecked then
        inc(Finish);
      tbMatrix[Finish, 0] := 'Scaled Total';
      tbMatrix[Finish, 1] := formatfloat('###', FTestStatsArray[1, 1]);
      for j := 2 to 5 do
        tbMatrix[Finish, j] := formatfloat(prec, FTestStatsArray[4, j]);
      for j := 6 to 7 + coladj do
        tbMatrix[Finish, j] := '-';
    end;

    if FNumDomains <> 1 then
    begin
      Domain := 0;
      if ScaledDomain.IsChecked then
        for row := Finish + 1 to summarytablerows - 1 do
        begin
          Domain := Domain + 1;
          tbMatrix[row, 0] := 'Scaled ' + FStoreDomain[Domain];
          tbMatrix[row, 1] := formatfloat('###', FDomainStatsArray[Domain, 1]);
          for j := 2 to 3 do
            tbMatrix[row, j] := formatfloat(prec, FDomScaleStat[Domain, j - 1]);
          for j := 4 to 5 do
            tbMatrix[row, j] := formatfloat(prec, FDomScaleStat[Domain, j - 1]);
          for j := 6 to 7 + coladj do
            tbMatrix[row, j] := '-';
        end;
    end;

    // For row := 0 to SummaryTableRows-1 do
    // begin
    // tbMatrix[row,0].bestwidth := 110;
    // tbMatrix[row,1].bestwidth := 45;
    // for j := 2 to 7 do SummaryTable.cells[row,j].bestwidth := 55;
    // if coladj = 1 then SummaryTable.cells[row,8].bestwidth := 60;
    // end;

    AddTable('Table 2: Summary statistics', tbMatrix);

    summarytablerows := 2;
    if FPreTest > 0 then
      summarytablerows := summarytablerows + 2;
    if FNumDomains > 1 then
      summarytablerows := summarytablerows + FNumDomains;
    { Write reliability analysis to rtf output AS A RICHVIEW TABLE }
    SetLength(tbMatrix, 0);
    SetLength(tbMatrix, summarytablerows, 9);

    tbMatrix[0, 0] := 'Score';
    tbMatrix[0, 1] := 'Alpha';
    tbMatrix[0, 2] := 'SEM';
    tbMatrix[0, 3] := 'Split-Half (Random)';
    tbMatrix[0, 4] := 'Split-Half (First-Last)';
    tbMatrix[0, 5] := 'Split-Half (Odd-Even)';
    tbMatrix[0, 6] := 'S-B Random';
    tbMatrix[0, 7] := 'S-B First-Last';
    tbMatrix[0, 8] := 'S-B Odd-Even';
    { TODO : tABLE COLOR }
    // for j := 0 to 8 do
    // tbMatrix[0,j].color := rgb(198,217,241);

    i := 0;
    if FPreTest > 0 then
    begin
      tbMatrix[1, 0] := 'All items';
      tbMatrix[1, 1] := formatfloat(prec, FTestStatsArray[2, 6]);
      tbMatrix[1, 2] := formatfloat(prec, FTestStatsArray[2, 9]);
      tbMatrix[1, 3] := formatfloat(prec, FTestStatsArray[2, 10]);
      tbMatrix[1, 4] := formatfloat(prec, FTestStatsArray[2, 11]);
      tbMatrix[1, 5] := formatfloat(prec, FTestStatsArray[2, 12]);
      tbMatrix[1, 6] := formatfloat(prec, FTestStatsArray[2, 13]);
      tbMatrix[1, 7] := formatfloat(prec, FTestStatsArray[2, 14]);
      tbMatrix[1, 8] := formatfloat(prec, FTestStatsArray[2, 15]);
      i := 1;
    end;
    v := 0;
    if FPreTest = 0 then
      v := 1;
    tbMatrix[1 + i, 0] := 'Scored items';
    tbMatrix[1 + i, 1] := formatfloat(prec, FTestStatsArray[1, 6]);
    tbMatrix[1 + i, 2] := formatfloat(prec, FTestStatsArray[1, 9]);
    tbMatrix[1 + i, 3] := formatfloat(prec, FTestStatsArray[1 + v, 10]);
    tbMatrix[1 + i, 4] := formatfloat(prec, FTestStatsArray[1 + v, 11]);
    tbMatrix[1 + i, 5] := formatfloat(prec, FTestStatsArray[1 + v, 12]);
    tbMatrix[1 + i, 6] := formatfloat(prec, FTestStatsArray[1 + v, 13]);
    tbMatrix[1 + i, 7] := formatfloat(prec, FTestStatsArray[1 + v, 14]);
    tbMatrix[1 + i, 8] := formatfloat(prec, FTestStatsArray[1 + v, 15]);

    if FPreTest > 0 then
    begin
      tbMatrix[3, 0] := 'Pretest items';
      tbMatrix[3, 1] := formatfloat(prec, FTestStatsArray[3, 6]);
      tbMatrix[3, 2] := formatfloat(prec, FTestStatsArray[3, 9]);
      for j := 3 to 6 do
        tbMatrix[3, j] := '-';
      i := 2;
    end;

    if FNumDomains > 1 then
      for h := 1 to FNumDomains do
      begin
        tbMatrix[1 + i + h, 0] := FStoreDomain[h];
        tbMatrix[1 + i + h, 1] := formatfloat(prec, FDomainStatsArray[h, 6]);
        if FNumberOfExaminees > 1 then
          if FDomainStatsArray[h, 6] <= 1 then
            FDomainStatsArray[h, 9] := FDomainStatsArray[h, 3] * sqrt(1 - FDomainStatsArray[h, 6]);

        tbMatrix[1 + i + h, 2] := formatfloat(prec, FDomainStatsArray[h, 9]);
        for j := 3 to 8 do
          tbMatrix[1 + i + h, j] := formatfloat(prec, FDomainStatsArray[h, 7 + j]);
      end;

    s := 'Table 3 presents a reliability analysis of the tests.  Alpha (also known as KR-20)';
    s := s + ' is the most commonly used index of reliability, and is therefore used to calculate';
    s := s + ' the standard error of measurement (SEM) on the raw score scale.  Also presented';
    s := s + ' are three configurations of split-half reliability, first as uncorrected';
    s := s + ' correlations, and then as Spearman-Brown (S-B) corrected correlations.';
    s := s + '  This is because an uncorrected split-half correlation is referenced to a "test" that';
    s := s + ' only contains half as many items as the full test, and therefore underestimates reliability.';
    AddText(PWideChar(s));
    // FExport.AddText('');
    AddText('');
    if (classcheckbox.IsChecked) then
    begin
      s := 'The cutscore on this exam was ' + classcutbox.Text + ', producing a pass rate of ' +
        formatfloat('0.0', 100 * FProppassing) + '%.';
      if not scoredrate then
        s := s + '  The Livingston index of classification consistency at the cut-score was ' +
          formatfloat(prec, Livingston()) + '.';
      AddText(PWideChar(s));
      AddText('');
    end;

    AddTable('Table 3: Reliability', tbMatrix);

    { Write specs to docx output }
    IF FNumFlagItems > 0 THEN
    BEGIN
      SetLength(tbMatrix, 0);
      SetLength(tbMatrix, FNumFlagItems + 1, 4);

      tbMatrix[0, 0] := 'Item ID';
      tbMatrix[0, 1] := 'P / Item Mean';
      tbMatrix[0, 2] := 'R';
      tbMatrix[0, 3] := 'Flag(s)';

      flagrow := 0;
      for row := 1 to FNumberOfItems do
        if row = FFlagID[row] then
        begin
          inc(flagrow);
          tbMatrix[flagrow, 0] := itemidarray[row];
          tbMatrix[flagrow, 1] := formatfloat('0.000', FItemStatsArray[row, 2]);
          tbMatrix[flagrow, 2] := formatfloat('0.000', FItemStatsArray[row, 3]);
          tbMatrix[flagrow, 3] := '';
          checkkey := False;
          { Flags }
          for j := 1 to itemscalearray[row, 1] do
            if (iteminfoarray[row, 3] = 'M') and (FItemStatsArray[row, 6] <> 0) then
              if FItemStatsArray[row, 3] < FOptionRpbisArray[row, j] then
                checkkey := True;
          if checkkey = True then
            tbMatrix[flagrow, 3] := FFlagArray[1];
          if iteminfoarray[row, 3] = 'M' then
          begin
            if checkkey = True then
              PreStr := ', '
            else
              PreStr := '';
            if (FItemStatsArray[row, 2] < FMinP) or (FItemStatsArray[row, 2] > FMaxP) then
              tbMatrix[flagrow, 3] := tbMatrix[flagrow, 3] + PreStr;
            If FItemStatsArray[row, 2] < FMinP then
              tbMatrix[flagrow, 3] := tbMatrix[flagrow, 3] + FFlagArray[2];
            If FItemStatsArray[row, 2] > FMaxP then
              tbMatrix[flagrow, 3] := tbMatrix[flagrow, 3] + FFlagArray[3];
            if (FItemStatsArray[row, 2] < FMinP) or (FItemStatsArray[row, 2] > FMaxP) then
              checkkey := True;
            if (FItemStatsArray[row, 2] < FMinP) or (FItemStatsArray[row, 2] > FMaxP) then
              PreStr := ', '
            else if checkkey = True then
              PreStr := ', '
            else
              PreStr := '';
          end;

          if iteminfoarray[row, 3] <> 'M' then
          begin
            if checkkey = True then
              PreStr := ', '
            else
              PreStr := '';
            if (FItemStatsArray[row, 2] < FMinMean) or (FItemStatsArray[row, 2] > FMaxMean) then
              tbMatrix[flagrow, 3] := tbMatrix[flagrow, 3] + PreStr;
            If FItemStatsArray[row, 2] < FMinMean then
              tbMatrix[flagrow, 3] := tbMatrix[flagrow, 3] + FFlagArray[6];
            If FItemStatsArray[row, 2] > FMaxMean then
              tbMatrix[flagrow, 3] := tbMatrix[flagrow, 3] + FFlagArray[7];
            if (FItemStatsArray[row, 2] < FMinMean) or (FItemStatsArray[row, 2] > FMaxMean) then
              checkkey := True;
            if (FItemStatsArray[row, 2] < FMinMean) or (FItemStatsArray[row, 2] > FMaxMean) then
              PreStr := ', '
            else if checkkey = True then
              PreStr := ', '
            else
              PreStr := '';
          end;

          if (FItemStatsArray[row, 3] < FMinR) or (FItemStatsArray[row, 3] > FMaxR) then
            tbMatrix[flagrow, 3] := tbMatrix[flagrow, 3] + PreStr;
          If FItemStatsArray[row, 3] < FMinR then
            tbMatrix[flagrow, 3] := tbMatrix[flagrow, 3] + FFlagArray[4];
          If FItemStatsArray[row, 3] > FMaxR then
            tbMatrix[flagrow, 3] := tbMatrix[flagrow, 3] + FFlagArray[5];
          if (FItemStatsArray[row, 3] < FMinR) or (FItemStatsArray[row, 3] > FMaxR) then
            checkkey := True;
          if (FItemStatsArray[row, 3] < FMinR) or (FItemStatsArray[row, 3] > FMaxR) then
            PreStr := ', '
          else if checkkey = True then
            PreStr := ', '
          else
            PreStr := '';

          if (difcheckbox.IsChecked) and (iteminfoarray[row, 2] = 'Y') and (iteminfoarray[row, 9] = 'D') then
          begin
            if difarray[row, 5] < 0.05 then
              tbMatrix[flagrow, 3] := tbMatrix[flagrow, 3] + PreStr;
            if difarray[row, 5] < 0.05 then
              tbMatrix[flagrow, 3] := tbMatrix[flagrow, 3] + FFlagArray[8];
          end;

        end; // flagging

      s := 'Table 4 presents the item statistics and flags for the item(s) that were flagged during the analysis';
      // FExport.AddText('');
      AddText('');
      AddText(PWideChar(s));
      AddText('');
      AddTable('Table 4: Summary Statistics for the Flagged Items', tbMatrix);

    END // if 1 item is flagged
    else
    begin
      AddText('');
      AddText('No items were flagged during this analysis.');
    end; // if no items flagged
  end;
end;

procedure TItemanCalculation.PrintFreq10(const prec: string; Min, Max: Real);
var
  i, low1, high1: Integer;
  lowreal, highreal: Real;
  overrule, smallrange, userange: Boolean;
  tbMatrix: TTableMatrix;
begin

  overrule := False;
  smallrange := False;
  SetLength(tbMatrix, 11, 2);

  tbMatrix[0, 1] := 'Frequency';

  if Min = -1 then
    userange := True
  else
    userange := False;

  if Max - Min < 20 then
    overrule := True
  else
    overrule := False;

  if (overrule = True) or (userange = True) then
    tbMatrix[0, 0] := 'Score'
  else
    tbMatrix[0, 0] := 'Range';

  for i := 1 to nbin do
  begin
    if (overrule = True) or (userange = True) then
      lowreal := GroupedFreqArray[i, 1];
    if (overrule = True) or (userange = True) then
      highreal := GroupedFreqArray[i, 2];
    if (overrule = False) and (userange = False) then
    begin
      low1 := trunc(GroupedFreqArray[i, 1]);
      high1 := trunc(GroupedFreqArray[i, 2]);
      if (GroupedFreqArray[i, 2] = high1) and (i <> nbin) then
        high1 := high1 - 1; // because last bin always includes max
      if (low1 < GroupedFreqArray[i, 1]) and (i <> 1) then
        low1 := low1 + 1; // because 1st bin always includes minimum
      lowreal := low1;
      highreal := high1;
    end;
    if (highreal - lowreal) < 1E-9 then
      smallrange := True
    else
      smallrange := False;
    if ((overrule) or (smallrange)) and (userange = False) then
      tbMatrix[i, 0] := formatfloat(prec, lowreal);
    if (userange = True) or ((overrule = False) and (smallrange = False)) then
      tbMatrix[i, 0] := formatfloat(prec, lowreal) + ' to ' + formatfloat(prec, highreal);
    tbMatrix[i, 1] := formatfloat('0', GroupedFreqArray[i, 3]);
  end;

  AddTable('', tbMatrix);
end;

procedure TItemanCalculation.PrintFreq15(const prec: string; Min, Max: Real);
var
  i, low1, high1: Integer;
  lowreal, highreal: Real;
  overrule, smallrange, userange: Boolean;
  tbMatrix: TTableMatrix;
begin
  overrule := False;
  smallrange := False;
  SetLength(tbMatrix, 16, 2);

  if Min = -1 then
    userange := True
  else
    userange := False;
  if Max - Min < 20 then
    overrule := True
  else
    overrule := False;

  if (overrule = True) or (userange = True) then
    tbMatrix[0, 0] := 'Score'
  else
    tbMatrix[0, 0] := 'Range';
  tbMatrix[0, 1] := 'Frequency';

  for i := 1 to nbin do
  begin
    if (overrule = True) or (userange = True) then
      lowreal := GroupedFreqArray[i, 1];
    if (overrule = True) or (userange = True) then
      highreal := GroupedFreqArray[i, 2];
    if (overrule = False) and (userange = False) then
    begin
      low1 := trunc(GroupedFreqArray[i, 1]);
      high1 := trunc(GroupedFreqArray[i, 2]);
      if (GroupedFreqArray[i, 2] = high1) and (i <> nbin) then
        high1 := high1 - 1; // because last bin always includes max
      if (low1 < GroupedFreqArray[i, 1]) and (i <> 1) then
        low1 := low1 + 1; // because 1st bin always includes minimum
      lowreal := low1;
      highreal := high1;
    end;
    if (highreal - lowreal) < 1E-9 then
      smallrange := True
    else
      smallrange := False;
    if ((overrule) or (smallrange)) and (userange = False) then
      tbMatrix[i, 0] := formatfloat(prec, lowreal);
    if (userange = True) or ((overrule = False) and (smallrange = False)) then
      tbMatrix[i, 0] := formatfloat(prec, lowreal) + ' to ' + formatfloat(prec, highreal);

    tbMatrix[i, 1] := formatfloat('0', GroupedFreqArray[i, 3]);
  end;

  AddTable('', tbMatrix);
end;

procedure TItemanCalculation.PrintFreq20(const prec: string; Min, Max: Real);
var
  i, low1, high1: Integer;
  lowreal, highreal: Real;
  overrule, smallrange, userange: Boolean;
  tbMatrix: TTableMatrix;
begin

  overrule := False;
  smallrange := False;
  SetLength(tbMatrix, 22, 2);

  if Min = -1 then
    userange := True
  else
    userange := False;
  if Max - Min < 20 then
    overrule := True
  else
    overrule := False;
  if (overrule = True) or (userange = True) then
    tbMatrix[0, 0] := 'Score'
  else
    tbMatrix[0, 0] := 'Range';
  tbMatrix[0, 1] := 'Frequency';

  for i := 1 to nbin do
  begin
    if (overrule = True) or (userange = True) then
      lowreal := GroupedFreqArray[i, 1];
    if (overrule = True) or (userange = True) then
      highreal := GroupedFreqArray[i, 2];
    if (overrule = False) and (userange = False) then
    begin
      low1 := trunc(GroupedFreqArray[i, 1]);
      high1 := trunc(GroupedFreqArray[i, 2]);
      if (GroupedFreqArray[i, 2] = high1) and (i <> nbin) then
        high1 := high1 - 1; // because last bin always includes max
      if (low1 < GroupedFreqArray[i, 1]) and (i <> 1) then
        low1 := low1 + 1; // because 1st bin always includes minimum
      lowreal := low1;
      highreal := high1;
    end;
    if (highreal - lowreal) < 1E-9 then
      smallrange := True
    else
      smallrange := False;
    if ((overrule) or (smallrange)) and (userange = False) then
      tbMatrix[i, 0] := formatfloat(prec, lowreal);
    if (userange = True) or ((overrule = False) and (smallrange = False)) then
      tbMatrix[i, 0] := formatfloat(prec, lowreal) + ' to ' + formatfloat(prec, highreal);
    tbMatrix[i, 1] := formatfloat('0', GroupedFreqArray[i, 3]);
  end;
  AddTable('', tbMatrix);
end;

procedure TItemanCalculation.PrintItemTables(const prec: string; item: Integer; iteminfoarray: TChar2Array;
  itemidarray: TStringArray; itemscalearray, nkeyarray: TIntegerArray; start0: Char; difarray: TRealArray);
var
  checkkey: Boolean;
  PreStr, PreText: string;
  DIFInt, j, d, g, REO, i, k, SCO, row: Integer;
  // REO is reorder variable if polytomous codes are reversed.
  tbMatrix: TTableMatrix;
begin
  REO := 0;
  with frmMain do
  begin
    if plotsbox.IsChecked = False then
      i := 1
    else
      i := 0; // no colors needed in tables
    if FNumDomains > 1 then
      g := 2
    else
      g := 0;

    { Information Table }
    SetLength(tbMatrix, 2, 7);

    tbMatrix[0, 0] := 'Seq.';
    tbMatrix[0, 1] := 'ID';
    tbMatrix[0, 2] := 'Key';
    tbMatrix[0, 3] := 'Scored';
    tbMatrix[0, 4] := 'Num Options';
    tbMatrix[0, 5] := 'Domain';
    tbMatrix[0, 6] := 'Flags';

    { Put in this item's results }
    if iteminfoarray[item, 2] = 'P' then
      PreText := 'Pretest'
    else
      PreText := 'Yes';
    tbMatrix[1, 0] := IntToStr(item);
    tbMatrix[1, 1] := itemidarray[item];
    if iteminfoarray[item, 1] <> '-' then
      tbMatrix[1, 2] := iteminfoarray[item, 1]
    else
      tbMatrix[1, 2] := 'Rev';
    if nkeyarray[item, 1] > 1 then
      for d := 1 to nkeyarray[item, 1] - 1 do
        tbMatrix[1, 2] := ',' + iteminfoarray[item, 3 + d];
    tbMatrix[1, 3] := PreText;
    tbMatrix[1, 4] := IntToStr(itemscalearray[item, 1]);
    tbMatrix[1, 5] := FStoreDomain[itemscalearray[item, 2]];

    checkkey := False;

    { Flags }
    for j := 1 to itemscalearray[item, 1] do
      if (iteminfoarray[item, 3] = 'M') and (FItemStatsArray[item, 6] <> 0) then
        if FItemStatsArray[item, 3] < FOptionRpbisArray[item, j] then
          checkkey := True;
    if checkkey = True then
      tbMatrix[1, 6] := FFlagArray[1];

    if iteminfoarray[item, 3] = 'M' then
    begin
      if checkkey = True then
        PreStr := ', '
      else
        PreStr := '';
      if (FItemStatsArray[item, 2] < FMinP) or (FItemStatsArray[item, 2] > FMaxP) then
        tbMatrix[1, 6] := PreStr;
      If FItemStatsArray[item, 2] < FMinP then
        tbMatrix[1, 6] := FFlagArray[2];
      If FItemStatsArray[item, 2] > FMaxP then
        tbMatrix[1, 6] := FFlagArray[3];
      if (FItemStatsArray[item, 2] < FMinP) or (FItemStatsArray[item, 2] > FMaxP) then
        checkkey := True;
      if (FItemStatsArray[item, 2] < FMinP) or (FItemStatsArray[item, 2] > FMaxP) then
        PreStr := ', '
      else if checkkey = True then
        PreStr := ', '
      else
        PreStr := '';
    end;

    if iteminfoarray[item, 3] <> 'M' then
    begin
      if checkkey = True then
        PreStr := ', '
      else
        PreStr := '';
      if (FItemStatsArray[item, 2] < FMinMean) or (FItemStatsArray[item, 2] > FMaxMean) then
        tbMatrix[1, 6] := PreStr;
      If FItemStatsArray[item, 2] < FMinMean then
        tbMatrix[1, 6] := FFlagArray[6];
      If FItemStatsArray[item, 2] > FMaxMean then
        tbMatrix[1, 6] := FFlagArray[7];
      if (FItemStatsArray[item, 2] < FMinMean) or (FItemStatsArray[item, 2] > FMaxMean) then
        checkkey := True;
      if (FItemStatsArray[item, 2] < FMinMean) or (FItemStatsArray[item, 2] > FMaxMean) then
        PreStr := ', '
      else if checkkey = True then
        PreStr := ', '
      else
        PreStr := '';
    end;

    if (FItemStatsArray[item, 3] < FMinR) or (FItemStatsArray[item, 3] > FMaxR) then
      tbMatrix[1, 6] := PreStr;
    if (FItemStatsArray[item, 3] < FMinR) or (FItemStatsArray[item, 3] > FMaxR) then
      checkkey := True;
    if (FItemStatsArray[item, 3] < FMinR) or (FItemStatsArray[item, 3] > FMaxR) then
      PreStr := ', '
    else if checkkey = True then
      PreStr := ', '
    else
      PreStr := '';
    If FItemStatsArray[item, 3] < FMinR then
      tbMatrix[1, 6] := FFlagArray[4];
    If FItemStatsArray[item, 3] > FMaxR then
      tbMatrix[1, 6] := FFlagArray[5];

    if (difcheckbox.IsChecked) and (iteminfoarray[item, 2] = 'Y') and (iteminfoarray[item, 9] = 'D') then
    begin
      if difarray[item, 5] < 0.05 then
        tbMatrix[1, 6] := PreStr;
      if difarray[item, 5] < 0.05 then
        tbMatrix[1, 6] := FFlagArray[8];
    end;

    AddTable('Item information', tbMatrix);

    { Item Statistics Table }
    if difcheckbox.IsChecked then
      DIFInt := 3
    else
      DIFInt := 0;
    { Create item stat table to be used for each item }
    SetLength(tbMatrix, 0);
    SetLength(tbMatrix, 2, 5 + g + DIFInt);
    tbMatrix[0, 0] := 'N';

    if iteminfoarray[item, 3] = 'M' then
      tbMatrix[0, 1] := 'P'
    else
      tbMatrix[0, 1] := 'Mean';
    if FNumDomains > 1 then
    begin
      if iteminfoarray[item, 3] <> 'M' then
        tbMatrix[0, 2] := 'Domain R'
      else
        tbMatrix[0, 2] := 'Domain Rpbis';
      if (iteminfoarray[item, 3] <> 'M') and (itemscalearray[item, 1] > 2) then
        tbMatrix[0, 3] := 'Domain Eta'
      else
        tbMatrix[0, 3] := 'Domain Rbis';
    end;

    if iteminfoarray[item, 3] <> 'M' then
      tbMatrix[0, 2 + g] := 'Total R'
    else
      tbMatrix[0, 2 + g] := 'Total Rpbis';
    if (iteminfoarray[item, 3] <> 'M') and (itemscalearray[item, 1] > 2) then
      tbMatrix[0, 3 + g] := 'Total Eta'
    else
      tbMatrix[0, 3 + g] := 'Total Rbis';
    tbMatrix[0, 4 + g] := 'Alpha w/o';

    { Put in this item's results }
    tbMatrix[1, 0] := formatfloat('###', FItemStatsArray[item, 1]);
    tbMatrix[1, 1] := formatfloat(prec, FItemStatsArray[item, 2]);
    if FNumDomains > 1 then
    begin
      tbMatrix[1, 2] := formatfloat(prec, FItemStatsArray[item, 7]);
      // domain rpbis/R
      tbMatrix[1, 3] := formatfloat(prec, FItemStatsArray[item, 8]);
      // domain rbis/eta
    end;
    tbMatrix[1, 2 + g] := formatfloat(prec, FItemStatsArray[item, 3]);
    tbMatrix[1, 3 + g] := formatfloat(prec, FItemStatsArray[item, 4]);
    if iteminfoarray[item, 2] = 'Y' then
      PreStr := formatfloat(prec, FItemStatsArray[item, 5])
    else
      PreStr := 'NA';
    tbMatrix[1, 4 + g] := PreStr;

    if difcheckbox.IsChecked then
    begin
      tbMatrix[0, 5 + g] := 'M-H';
      tbMatrix[0, 6 + g] := 'p';
      tbMatrix[0, 7 + g] := 'Bias Against';
      tbMatrix[1, 5 + g] := formatfloat(prec, difarray[item, 1]);
      tbMatrix[1, 6 + g] := formatfloat(prec, difarray[item, 5]);
      if (difarray[item, 5] < 0.05) and (difarray[item, 4] < 0) then
        tbMatrix[1, 7 + g] := Group2label.Text;
      if (difarray[item, 5] < 0.05) and (difarray[item, 4] > 0) then
        tbMatrix[1, 7 + g] := Group1label.Text;
      if difarray[item, 5] > 0.05 then
        tbMatrix[1, 7 + g] := 'N/A';
    end;

    AddTable('Item statistics', tbMatrix);

    { Option stats table }
    { Create item stat table to be used for each item }
    SetLength(tbMatrix, 0);
    SetLength(tbMatrix, FMaxItemOptions + 3, 9 - i);

    tbMatrix[0, 0] := 'Option';
    if iteminfoarray[item, 3] <> 'M' then
      j := 1
    else
      j := 0;
    if iteminfoarray[item, 3] <> 'M' then
      tbMatrix[0, 1] := 'Weight';
    tbMatrix[0, 1 + j] := 'N';
    tbMatrix[0, 2 + j] := 'Prop.';
    tbMatrix[0, 3 + j] := 'Rpbis';
    tbMatrix[0, 4 + j] := 'Rbis';
    tbMatrix[0, 5 + j] := 'Mean';
    tbMatrix[0, 6 + j] := 'SD';
    if plotsbox.IsChecked then
      tbMatrix[0, 7 + j] := 'Color';

    For row := 1 to itemscalearray[item, 1] + 2 do
    begin
      if row <= itemscalearray[item, 1] then
      begin
        if (FRespAsInt = True) and (start0 <> 'P') then
          tbMatrix[row, 0] := IntToStr(row);
        if start0 = 'P' then
          tbMatrix[row, 0] := IntToStr(row - 1);
        if (FRespAsInt = False) and (start0 <> 'P') then
          tbMatrix[row, 0] := chr(row + 64);
        j := 0;
        if iteminfoarray[item, 3] <> 'M' then
        begin
          j := 1;
          REO := row;
          if iteminfoarray[item, 1] = '+' then
            if iteminfoarray[item, 3] = 'P' then
              REO := row - 1;
          if (iteminfoarray[item, 1] <> '+') and (iteminfoarray[item, 1] <> '-') then
            if iteminfoarray[item, 3] = 'P' then
              REO := row - 1;
          if iteminfoarray[item, 1] = '-' then
            if iteminfoarray[item, 3] = 'R' then
              REO := (itemscalearray[item, 1] + 1 - row)
            else
              REO := row;
          if iteminfoarray[item, 1] = '-' then
            if iteminfoarray[item, 3] = 'P' then
              REO := (itemscalearray[item, 1] - row);
          tbMatrix[row, 1] := IntToStr(REO);
          if iteminfoarray[item, 3] = 'P' then
            REO := REO + 1;
        end; // if item is rating scale or 'R'
        if iteminfoarray[item, 3] = 'M' then
          REO := row;
        tbMatrix[row, 1 + j] := IntToStr(FOptionCountsArray[item, REO]);
        If FOptionCountsArray[item, row] > 0 then
        begin
          tbMatrix[row, 2 + j] := formatfloat(prec, FOptionPropsArray[item, REO]);
          tbMatrix[row, 3 + j] := formatfloat(prec, FOptionRpbisArray[item, REO]);
          tbMatrix[row, 4 + j] := formatfloat(prec, FOptionRbisArray[item, REO]);
          tbMatrix[row, 5 + j] := formatfloat(prec, FOptionMeansArray[item, REO]);
          tbMatrix[row, 6 + j] := formatfloat(prec, FOptionSDArray[item, REO]);
        end
        else
        begin
          tbMatrix[row, 2 + j] := prec;
          tbMatrix[row, 3 + j] := '--';
          tbMatrix[row, 4 + j] := '--';
          tbMatrix[row, 5 + j] := '--';
          tbMatrix[row, 6 + j] := '--';
        end;
      end;
      if row = itemscalearray[item, 1] + 1 then
      begin
        tbMatrix[row, 0] := 'Omit';
        tbMatrix[row, 1 + j] := IntToStr(FOptionCountsArray[item, 16]);
        if (OmitAsinc.IsChecked = False) and (FOptionCountsArray[item, 16] > 0) then
          tbMatrix[row, 2 + j] := formatfloat(prec, FOptionPropsArray[item, 16]);
        if (OmitAsinc.IsChecked = False) and (FOptionCountsArray[item, 16] > 0) then
          tbMatrix[row, 3 + j] := formatfloat(prec, FOptionRpbisArray[item, 16]);
        if (OmitAsinc.IsChecked = False) and (FOptionCountsArray[item, 16] > 0) then
          tbMatrix[row, 4 + j] := formatfloat(prec, FOptionRbisArray[item, 16]);
        if FOptionCountsArray[item, 16] > 0 then
          tbMatrix[row, 5 + j] := formatfloat(prec, FOptionMeansArray[item, 16]);
        if FOptionCountsArray[item, 16] > 0 then
          tbMatrix[row, 6 + j] := formatfloat(prec, FOptionSDArray[item, 16]);
      end;
      if row = itemscalearray[item, 1] + 2 then
      begin
        tbMatrix[row, 0] := 'Not Admin';
        tbMatrix[row, 1 + j] := IntToStr(FOptionCountsArray[item, 17]);
        if FOptionCountsArray[item, 17] > 0 then
          tbMatrix[row, 5 + j] := formatfloat(prec, FOptionMeansArray[item, 17]);
        if FOptionCountsArray[item, 17] > 0 then
          tbMatrix[row, 6 + j] := formatfloat(prec, FOptionSDArray[item, 17]);
      end;
      if start0 = 'P' then
        SCO := 1
      else
        SCO := 0;
      If (ConvertResponseToInteger(iteminfoarray[item, 1], iteminfoarray, item, itemscalearray[item, 1]) = row - SCO)
        and (iteminfoarray[item, 3] = 'M') then
        tbMatrix[row, 8 - i] := '**KEY**';
      if nkeyarray[item, 1] > 1 then
        for k := 1 to nkeyarray[item, 1] - 1 do
          If (ConvertResponseToInteger(iteminfoarray[item, 3 + k], iteminfoarray, item, itemscalearray[item, 1]) = row -
            SCO) and (iteminfoarray[item, 3] = 'M') then
            tbMatrix[row, 8 - i] := '**KEY' + IntToStr(k + 1) + '**';
    end; // for row to itemscalearray + 2

    For row := 1 to itemscalearray[item, 1] do
      if plotsbox.IsChecked then
      begin // if no plotting, then colors not needed
        If row = 1 then
          tbMatrix[row, 7 + j] := 'Rust';
        If row = 2 then
          tbMatrix[row, 7 + j] := 'Navy';
        If row = 3 then
          tbMatrix[row, 7 + j] := 'Gold';
        If row = 4 then
          tbMatrix[row, 7 + j] := 'Gray';
        If row = 5 then
          tbMatrix[row, 7 + j] := 'LightBlue';
        If row = 6 then
          tbMatrix[row, 7 + j] := 'Navy';
        If row = 7 then
          tbMatrix[row, 7 + j] := 'Red';
        If row = 8 then
          tbMatrix[row, 7 + j] := 'Teal';
        If row = 9 then
          tbMatrix[row, 7 + j] := 'Purple';
      end; // for row loop

    AddTable('Option statistics', tbMatrix);
    if not tablebox.IsChecked then
      Exit;

    { Quantile table }
    SetLength(tbMatrix, 0);
    SetLength(tbMatrix, FMaxItemOptions + 1, 4 + FNumCut - i);

    { Create item stat table to be used for each item }

    tbMatrix[0, 0] := 'Option';
    tbMatrix[0, 1] := 'N';
    for j := 1 to round(CutPoint.Text.ToDouble) do
      tbMatrix[0, j + 1] := formatfloat('0', ((j - 1) / FNumCut) * 100) + '-' +
        formatfloat('0', (j / FNumCut) * 100) + '%';

    if plotsbox.IsChecked then
      tbMatrix[0, 2 + FNumCut] := 'Color';

    For row := 1 to itemscalearray[item, 1] do
    begin
      if iteminfoarray[item, 1] = '-' then
        REO := (itemscalearray[item, 1] + 1) - row
      else
        REO := row;
      if (FRespAsInt = True) and (iteminfoarray[item, 3] <> 'P') then
        tbMatrix[row, 0] := IntToStr(row);
      if iteminfoarray[item, 3] = 'P' then
        tbMatrix[row, 0] := IntToStr(row - 1);
      if (FRespAsInt = False) and (start0 <> 'P') then
        tbMatrix[row, 0] := chr(row + 64);
      tbMatrix[row, 1] := IntToStr(FOptionCountsArray[item, REO]);
      if plotsbox.IsChecked then
      begin // if no plotting, then colors not needed
        If row = 1 then
          tbMatrix[row, 2 + FNumCut] := 'Rust';
        If row = 2 then
          tbMatrix[row, 2 + FNumCut] := 'Navy';
        If row = 3 then
          tbMatrix[row, 2 + FNumCut] := 'Gold';
        If row = 4 then
          tbMatrix[row, 2 + FNumCut] := 'Gray';
        If row = 5 then
          tbMatrix[row, 2 + FNumCut] := 'LightBlue';
        If row = 6 then
          tbMatrix[row, 2 + FNumCut] := 'Navy';
        If row = 7 then
          tbMatrix[row, 2 + FNumCut] := 'Red';
        If row = 8 then
          tbMatrix[row, 2 + FNumCut] := 'Teal';
        If row = 9 then
          tbMatrix[row, 2 + FNumCut] := 'Purple';
      end;

      for j := 1 to FNumCut do
        tbMatrix[row, j + 1] := formatfloat(prec, FSubGroupPArray[((REO - 1) * FNumCut + j)]);
      If (ConvertResponseToInteger(iteminfoarray[item, 1], iteminfoarray, item, itemscalearray[item, 1]) = row - SCO)
        and (iteminfoarray[item, 3] = 'M') then
        tbMatrix[row, FNumCut + 3 - i] := '**KEY**';
      if nkeyarray[item, 1] > 1 then
        for k := 1 to nkeyarray[item, 1] - 1 do
          If (ConvertResponseToInteger(iteminfoarray[item, 3 + k], iteminfoarray, item, itemscalearray[item, 1]) = row -
            SCO) and (iteminfoarray[item, 3] = 'M') then
            tbMatrix[row, FNumCut + 3 - i] := '**KEY' + IntToStr(k + 1) + '**';
    end;
  end;

  AddTable('Quantile plot data', tbMatrix);
end;

function TItemanCalculation.cdfni(p: single): single;
var
  deriv: single;
  change: single;
  count: Integer;
  z: single;
begin
  if p < 0.0000003 then
    Result := -5.0
  else if p > 0.9999997 then
    Result := 5.0
  else
  begin
    count := 0;
    z := 0.7;

    repeat
      inc(count);

      deriv := 0.398942 * exp(-z * z / 2.0);
      change := (cdfn(z) - p) / deriv;
      z := z - change;
    until (abs(change) < 0.00001) or (count = 20);

    if abs(change) > 0.001 then
      writeln('*** WARNING ***  CDFNI failed to converge -- P = ', p:7:4, ', Z = ', z:8:4, '.');
    Result := z;
  end;
end;

procedure TItemanCalculation.ScaleStats(iteminfoarray: TChar2Array);
var
  lptemp: Real;
  lpsum: array [1 .. 6] of Real;
  lRow, ifor: Integer;
begin
  { --- Calculate alpha--- }
  for ifor := 1 to 3 do
    lpsum[ifor] := 0;

  for lRow := 1 to FNumberOfItems do
  begin
    if iteminfoarray[lRow, 2] = 'Y' then { If item is scored }
      lpsum[1] := lpsum[1] + FItemStatsArray[lRow, 6];

    if iteminfoarray[lRow, 2] <> 'N' then
      lpsum[2] := lpsum[2] + FItemStatsArray[lRow, 6];

    if iteminfoarray[lRow, 2] = 'P' then { If item is pretest }
      lpsum[3] := lpsum[3] + FItemStatsArray[lRow, 6];
  end;

  If FScoredItems > 1 then // alpha for scored items
    FTestStatsArray[1, 6] := (FScoredItems / (FScoredItems - 1)) *
      (1 - (lpsum[1] / (FTestStatsArray[1, 3] * FTestStatsArray[1, 3] + 0.000000000001)))
  else
    FTestStatsArray[1, 6] := 0;

  If FValidItems > 1 then // alpha for all items
    FTestStatsArray[2, 6] := (FValidItems / (FValidItems - 1)) *
      (1 - (lpsum[2] / (FTestStatsArray[2, 3] * FTestStatsArray[2, 3] + 0.000000000001)))
  else
    FTestStatsArray[2, 6] := 0;

  if FTestStatsArray[3, 1] > 1 then // alpha for pretest items
    FTestStatsArray[3, 6] := (FTestStatsArray[3, 1] / (FTestStatsArray[3, 1] - 1)) *
      (1 - (lpsum[3] / (FTestStatsArray[3, 3] * FTestStatsArray[3, 3] + 0.000000000001)))
  else
    FTestStatsArray[3, 6] := 0;

  { Calculate mean P }
  for ifor := 1 to 6 do
    lpsum[ifor] := 0;

  for lRow := 1 to FNumberOfItems do
  begin
    if iteminfoarray[lRow, 3] = 'M' then
    begin
      if iteminfoarray[lRow, 2] = 'Y' then
      begin { Item is scored }
        lptemp := FItemStatsArray[lRow, 2];
        lpsum[1] := lpsum[1] + lptemp;
      end;

      if iteminfoarray[lRow, 2] <> 'N' then
      begin
        lptemp := FItemStatsArray[lRow, 2];
        lpsum[2] := lpsum[2] + lptemp;
      end;

      If iteminfoarray[lRow, 2] = 'P' then
      begin { Item is pretest }
        lptemp := FItemStatsArray[lRow, 2];
        lpsum[3] := lpsum[3] + lptemp;
      end;
    end; // if mult choice

    if iteminfoarray[lRow, 3] <> 'M' then
    begin
      if iteminfoarray[lRow, 2] = 'Y' then
      begin { Item is scored }
        lptemp := FItemStatsArray[lRow, 2];
        lpsum[4] := lpsum[4] + lptemp;
      end;

      if iteminfoarray[lRow, 2] <> 'N' then
      begin
        lptemp := FItemStatsArray[lRow, 2];
        lpsum[5] := lpsum[5] + lptemp;
      end;

      if iteminfoarray[lRow, 2] = 'P' then
      begin { Item is pretest }
        lptemp := FItemStatsArray[lRow, 2];
        lpsum[6] := lpsum[6] + lptemp;
      end;
    end; // if rating scale
  end;

  if FIsMult then
  begin
    FTestStatsArray[1, 7] := lpsum[1] / (FNumMult[1] + 0.000000000001);
    FTestStatsArray[2, 7] := lpsum[2] / (FNumMult[2] + 0.000000000001);
    FTestStatsArray[3, 7] := lpsum[3] / (FNumMult[3] + 0.000000000001);
  end;

  if FIsRate then
  begin
    FTestStatsArray[1, 16] := lpsum[4] / (FNumRate[1] + 0.000000000001);
    FTestStatsArray[2, 16] := lpsum[5] / (FNumRate[2] + 0.000000000001);
    FTestStatsArray[3, 16] := lpsum[6] / (FNumRate[3] + 0.000000000001);
  end;

  { Calculate mean Rpbis of scored items }
  for ifor := 1 to 3 do
    lpsum[ifor] := 0.0;

  for lRow := 1 to FNumberOfItems do
  begin
    if iteminfoarray[lRow, 2] = 'Y' then
    begin { Item is scored }
      lptemp := FItemStatsArray[lRow, 3];
      lpsum[1] := lpsum[1] + lptemp;
    end;

    if iteminfoarray[lRow, 2] <> 'N' then
    begin
      lptemp := FItemStatsArray[lRow, 3];
      lpsum[2] := lpsum[2] + lptemp;
    end;

    if iteminfoarray[lRow, 2] = 'P' then
    begin { Item is pretest }
      lptemp := FItemStatsArray[lRow, 3];
      lpsum[3] := lpsum[3] + lptemp;
    end;
  end;

  FTestStatsArray[1, 8] := lpsum[1] / (FScoredItems + 0.000000000001);
  FTestStatsArray[2, 8] := lpsum[2] / (FTestStatsArray[2, 1] + 0.000000000001);
  FTestStatsArray[3, 8] := lpsum[3] / (FTestStatsArray[3, 1] + 0.000000000001);
end;

procedure TItemanCalculation.EstimateKR20Without(responsearray, iteminfoarray: TChar2Array;
  itemscalearray, nkeyarray, scorearray: TIntegerArray);
var
  lScoreSum: Real;
  k, i, j, row: Integer;
begin
  for i := 1 to FNumberOfItems do
    if iteminfoarray[i, 2] = 'Y' then
    begin
      FPSum := 0.0;

      for row := 1 to FNumberOfItems do
        if (iteminfoarray[row, 2] = 'Y') and (i <> row) then
        begin
          FPTemp := FItemStatsArray[row, 2];
          FPSum := FPSum + FItemStatsArray[row, 6];
        end;

      { Find scores without that item }
      lScoreSum := 0;
      for row := 1 to FNumberOfExaminees do
      begin
        scorearray[row, 3] := scorearray[row, 1];
        FCh := responsearray[row, i];

        if (FCh = iteminfoarray[i, 1]) and (iteminfoarray[i, 3] = 'M') then { If the response is the key }
          scorearray[row, 3] := scorearray[row, 3] - 1;
        { Subtracts the 1/0 indicator }

        j := 1;

        if nkeyarray[i, 1] > 1 then
          for k := 1 to nkeyarray[i, 1] - 1 do
          begin
            if (FCh = iteminfoarray[i, 3 + j]) and (iteminfoarray[i, 3] = 'M') then { If the response is the key }
              scorearray[row, 3] := scorearray[row, 3] - 1;
            { Subtracts the 1/0 indicator }

            if (FCh = iteminfoarray[i, 3 + j]) and (iteminfoarray[i, 3] = 'M') then
              j := j + 1;
          end;

        if iteminfoarray[i, 3] <> 'M' then
          scorearray[row, 3] := scorearray[row, 3] - ConvertResponseToInteger(FCh, iteminfoarray, i,
            itemscalearray[i, 1]);

        { Add that corrected score to running sum for the total corrected mean }
        lScoreSum := lScoreSum + scorearray[row, 3];
      end;
      FScoreMean := lScoreSum / FNumberOfExaminees;

      { Find FVarianceForKR20Without, which is without the item }
      FScoreSS := 0;
      for row := 1 to FNumberOfExaminees do
        FScoreSS := FScoreSS + (scorearray[row, 3] - FScoreMean) * (scorearray[row, 3] - FScoreMean);

      if FNumberOfExaminees > 1 then
        FVarianceForKR20Without := FScoreSS / (FNumberOfExaminees - 1);

      FItemStatsArray[i, 5] := ((FNumberOfItems - 1) / (FNumberOfItems - 2)) *
        (1 - (FPSum / (FVarianceForKR20Without + 0.000000000001)));
      // need updated variance
    end; // if 'y'
end;

procedure TItemanCalculation.DomainStats(iteminfoarray, responsearray: TChar2Array; itemscalearray: TIntegerArray;
  domscalearray: TRealArray; nkeyarray: TIntegerArray);
var
  Domain, domain1, domain2, row, resp, k: Integer;
  rescaled, i, j, temp4, temp1, temp2, temp3, temp5, sstot, ssbet, itemscore: Real;
  domainmean1, domainmean2, domainsd1, domainsd2: Real;
  mi, mj, TSD, TMean: Real; // min and max for scaled domain scores!
  dommean: array [1 .. 16] of Real;
  domcount: array [1 .. 16] of Integer;
  correct: Boolean;
  lColumn: Integer;
  lTempInt: Integer;
  lDomainMean: Real;
  lDomainSD: Real;
  lDomainSS: Real;
  lDomainSum: Real;
begin
  if FNumDomains > 1 then
  begin
    { ------- Reset to 0 ------- }
    for Domain := 1 to FNumDomains do
      for row := 1 to FNumberOfExaminees do
        FDomainArray[row, Domain] := 0;

    { ------------- Calculate domain scores ---------- }
    for Domain := 1 to FNumDomains do
      for row := 1 to FNumberOfExaminees do
        for lColumn := 1 to FNumberOfItems do
        begin
          if row = 1 then
            frmMain.ProgressBar.Value := frmMain.ProgressBar.Value + 1;
          Application.ProcessMessages;

          if (iteminfoarray[lColumn, 2] = 'Y') and (itemscalearray[lColumn, 2] = Domain) then
          begin
            if iteminfoarray[lColumn, 3] = 'M' then
              for k := 1 to nkeyarray[lColumn, 1] do
              begin
                if k = 1 then
                  lTempInt := 0
                else
                  lTempInt := 2; // 4 so 2+2 gives lColumn 4

                if responsearray[row, lColumn] = iteminfoarray[lColumn, lTempInt + k] then { Correct }
                  FDomainArray[row, Domain] := FDomainArray[row, Domain] + 1;
              end; // for nkeyarray

            if (iteminfoarray[lColumn, 3] <> 'M') and (responsearray[row, lColumn] <> FNAChar) and
              (responsearray[row, lColumn] <> FOmitChar) then
              FDomainArray[row, Domain] := FDomainArray[row, Domain] +
                ConvertResponseToInteger(responsearray[row, lColumn], iteminfoarray, lColumn,
                itemscalearray[lColumn, 1]);
          end;
        end;

    { ------- Reset to 0 ------- }
    for Domain := 1 to 8 do
      for row := 1 to FNumDomains do
        FDomainStatsArray[row, Domain] := 0;

    { --- Calculate number of items in each domain --- }
    for row := 1 to FNumberOfItems do
      if iteminfoarray[row, 2] = 'Y' then { Item is scored }
        FDomainStatsArray[itemscalearray[row, 2], 1] := FDomainStatsArray[itemscalearray[row, 2], 1] + 1;

    { ****Cycle through domains and calculate stats**** }
    { NItems, Mean, SD, Min, Max, Alpha, Mean P, Mean Rpbis }
    for Domain := 1 to FNumDomains do
    begin
      { ---Calculate overall stats--- }
      lDomainSum := 0;
      lDomainSS := 0;

      for row := 1 to FNumberOfExaminees do
        lDomainSum := lDomainSum + FDomainArray[row, Domain];

      lDomainMean := lDomainSum / FNumberOfExaminees;

      for row := 1 to FNumberOfExaminees do
        lDomainSS := lDomainSS + (lDomainMean - FDomainArray[row, Domain]) * (lDomainMean - FDomainArray[row, Domain]);

      if FNumberOfExaminees > 1 then
        lDomainSD := sqrt(lDomainSS / (FNumberOfExaminees - 1));

      FDomainStatsArray[Domain, 2] := lDomainMean;
      FDomainStatsArray[Domain, 3] := lDomainSD;

      { --- Min and Max--- }
      i := 10000;
      j := 0;
      mi := 10000;
      mj := 0;
      lDomainSum := 0;
      lDomainSS := 0;
      rescaled := 0;

      for row := 1 to FNumberOfExaminees do
      begin
        if FDomainArray[row, Domain] < i then
          i := FDomainArray[row, Domain];
        if FDomainArray[row, Domain] > j then
          j := FDomainArray[row, Domain];

        if (frmMain.standscale.IsChecked) and (FDomainStatsArray[Domain, 3] > 0) then
        begin
          rescaled := FDomainArray[row, Domain] *
            (StrToFloatDef(frmMain.newsdbox.Text, 0) / FDomainStatsArray[Domain, 3]);
          rescaled := rescaled + (StrToFloatDef(frmMain.newmeanbox.Text, 0) - (StrToFloatDef(frmMain.newsdbox.Text,
            0) / FDomainStatsArray[Domain, 3]) * FDomainStatsArray[Domain, 2]);
        end;

        if frmMain.linearscale.IsChecked then
          rescaled := FDomainArray[row, Domain] * StrToFloatDef(frmMain.slopebox.Text, 0) +
            StrToFloatDef(frmMain.interceptbox.Text, 0);

        domscalearray[row, Domain] := rescaled;
        lDomainSum := lDomainSum + rescaled;

        if rescaled < mi then
          mi := rescaled;
        if rescaled > mj then
          mj := rescaled;
      end; // row numberofexaminee loop

      FDomainStatsArray[Domain, 4] := i;
      FDomainStatsArray[Domain, 5] := j;
      FDomScaleStat[Domain, 3] := mi; // scaled domain min
      FDomScaleStat[Domain, 4] := mj; // scaled domain max
      FDomScaleStat[Domain, 1] := lDomainSum / FNumberOfExaminees;

      for row := 1 to FNumberOfExaminees do
        lDomainSS := lDomainSS + (FDomScaleStat[Domain, 1] - domscalearray[row, Domain]) *
          (FDomScaleStat[Domain, 1] - domscalearray[row, Domain]);

      if FNumberOfExaminees > 1 then
        FDomScaleStat[Domain, 2] := sqrt(lDomainSS / (FNumberOfExaminees - 1));

      FDomainStatsArray[Domain, 6] := EstimateKR20(Domain, iteminfoarray, itemscalearray);
      FDomainStatsArray[Domain, 7] := CalculateMeanP(Domain, iteminfoarray, itemscalearray);
      FDomainStatsArray[Domain, 16] := CalculateItemMean(Domain, iteminfoarray, itemscalearray);
      FDomainStatsArray[Domain, 8] := CalculateMeanRpbis(Domain, iteminfoarray, itemscalearray);

      for lColumn := 1 to FNumberOfItems do
        if iteminfoarray[lColumn, 2] <> 'N' then
        begin
          temp1 := 0;
          temp2 := 0;
          temp4 := 0;
          TMean := 0;
          TSD := 0;
          FItemN := 0;
          ssbet := 0;
          frmMain.ProgressBar.Value := frmMain.ProgressBar.Value + 1;
          Application.ProcessMessages;

          for row := 1 to 16 do
            dommean[row] := 0; // reset to 0
          for row := 1 to 16 do
            domcount[row] := 0; // reset to 0

          if (iteminfoarray[lColumn, 3] <> 'M') and (itemscalearray[lColumn, 2] = Domain) then
          begin
            for row := 1 to FNumberOfExaminees do
              if (responsearray[row, lColumn] <> FOmitChar) and (responsearray[row, lColumn] <> FNAChar) then
              begin
                inc(FItemN);
                FCh := responsearray[row, lColumn];
                FResponseAsInt := ConvertResponseToInteger(FCh, iteminfoarray, lColumn, itemscalearray[lColumn, 1]);

                if not frmMain.spurbox.IsChecked then
                  FResponseAsInt := 0;

                TMean := (FDomainArray[row, Domain] - FResponseAsInt) + TMean;
              end;

            TMean := TMean / (FItemN + 1E-20);

            for row := 1 to FNumberOfExaminees do
              if (responsearray[row, lColumn] <> FOmitChar) and (responsearray[row, lColumn] <> FNAChar) then
              begin
                FCh := responsearray[row, lColumn];
                FResponseAsInt := ConvertResponseToInteger(FCh, iteminfoarray, lColumn, itemscalearray[lColumn, 1]);

                if (frmMain.spurbox.IsChecked = False) then
                  FResponseAsInt := 0;

                TSD := TSD + ((FDomainArray[row, Domain] - FResponseAsInt) - TMean) *
                  ((FDomainArray[row, Domain] - FResponseAsInt) - TMean);
              end;

            if FItemN > 1 then
              TSD := sqrt(TSD / (FItemN - 1));

            for row := 1 to FNumberOfExaminees do
            begin
              FCh := responsearray[row, lColumn];
              FResponseAsInt := ConvertResponseToInteger(FCh, iteminfoarray, lColumn, itemscalearray[lColumn, 1]);

              if (responsearray[row, lColumn] <> FOmitChar) and (responsearray[row, lColumn] <> FNAChar) then
              begin
                if (frmMain.spurbox.IsChecked = False) then
                  itemscore := 0
                else
                  itemscore := FResponseAsInt;

                if frmMain.spurbox.IsChecked then
                  temp4 := temp4 + (FDomainArray[row, Domain] - itemscore - TMean) *
                    (FDomainArray[row, Domain] - itemscore - TMean);

                dommean[FResponseAsInt + FWts] := dommean[FResponseAsInt + FWts] +
                  (FDomainArray[row, Domain] - itemscore);
                domcount[FResponseAsInt + FWts] := domcount[FResponseAsInt + FWts] + 1;
                temp1 := temp1 + (FResponseAsInt - FItemStatsArray[lColumn, 2]) *
                  (FResponseAsInt - FItemStatsArray[lColumn, 2]);
                temp2 := temp2 + (FResponseAsInt - FItemStatsArray[lColumn, 2]) *
                  (FDomainArray[row, Domain] - itemscore - TMean);
              end; // if response is not omit FNAChar
            end; // FNumberOfExaminees loop

            for row := 1 to itemscalearray[lColumn, 1] do
              dommean[row] := dommean[row] / (domcount[row] + 1E-20);

            if frmMain.spurbox.IsChecked then
              sstot := temp4
            else
              sstot := TSD * TSD * (FItemN - 1); // for eta

            if FItemN > 1 then
              temp4 := temp4 / (FItemN - 1 + 1E-20); // variance
            if FItemN > 1 then
              temp5 := temp1 / (FItemN - 1 + 1E-20);

            for row := 1 to itemscalearray[lColumn, 1] do
              ssbet := ssbet + domcount[row] * (dommean[row] - TMean) * (dommean[row] - TMean);

            if frmMain.spurbox.IsChecked then
              temp3 := power(temp5, 0.5) * power(temp4, 0.5) * (FItemN - 1)
            else
              temp3 := power(temp5, 0.5) * TSD * (FItemN - 1);

            FItemStatsArray[lColumn, 7] := temp2 / (temp3 + 1E-22);
            // domain R correlation for polytomous items
            FItemStatsArray[lColumn, 8] := sqrt(ssbet / (sstot + 1E-20));
          end; // if item is in the domain if statement

          FP := FItemStatsArray[lColumn, 2];

          if iteminfoarray[lColumn, 3] = 'M' then
            if itemscalearray[lColumn, 2] = Domain then
              for row := 1 to FNumberOfExaminees do
                if responsearray[row, lColumn] <> FNAChar then
                  if (responsearray[row, lColumn] <> FOmitChar) or (frmMain.OmitAsinc.IsChecked = False) then
                  begin
                    inc(FItemN);
                    correct := False;
                    FCh := responsearray[row, lColumn];

                    if FCh = iteminfoarray[lColumn, 1] then
                      correct := True;
                    if (nkeyarray[lColumn, 1] > 1) then
                      for k := 1 to nkeyarray[lColumn, 1] - 1 do
                        if UpCase(iteminfoarray[lColumn, k + 3]) = UpCase(FCh) then
                          correct := True;

                    if correct = True then
                      resp := 1
                    else
                      resp := 0;
                    if (resp = 1) and frmMain.spurbox.IsChecked then
                      temp4 := 1
                    else
                      temp4 := 0;
                    if (resp = 1) and frmMain.spurbox.IsChecked then
                      temp2 := FP
                    else
                      temp2 := 0;
                    if (FP <> 0) and (FP <> 1) and (FDomainStatsArray[Domain, 3] <> 0) then
                      temp1 := temp1 + (((FDomainArray[row, Domain] - temp4) - (FDomainStatsArray[Domain, 2] - temp2)) /
                        FDomainStatsArray[Domain, 3]) * (resp - FP) / sqrt(FP * (1 - FP) + 1E-20);

                    FItemStatsArray[lColumn, 7] := temp1 / (FItemN - 1 + 1E-20);

                    { Convert Rpbis to Rbis }
                    FZ := -cdfni(FP);
                    FOrdinate := 0.398942 * exp(-sqr(FZ) / 2);
                    FItemStatsArray[lColumn, 8] := FItemStatsArray[lColumn, 7] * (sqrt(FP * (1 - FP)) / FOrdinate);

                    if (iteminfoarray[lColumn, 3] = 'P') and (itemscalearray[lColumn, 1] = 2) then
                      FP := FItemStatsArray[lColumn, 2];
                    if (iteminfoarray[lColumn, 3] = 'R') and (itemscalearray[lColumn, 1] = 2) then
                      FP := FItemStatsArray[lColumn, 2] - 1;

                    if (iteminfoarray[lColumn, 3] <> 'M') and (itemscalearray[lColumn, 1] = 2) then
                    begin
                      FZ := -cdfni(FP);
                      FOrdinate := 0.398942 * exp(-sqr(FZ) / 2);
                      FItemStatsArray[lColumn, 8] := FItemStatsArray[lColumn, 7] * (sqrt(FP * (1 - FP)) / FOrdinate);
                    end;

                    if FItemStatsArray[lColumn, 8] > 1.0 then
                      FItemStatsArray[lColumn, 8] := 1.0; { cap }
                    if FItemStatsArray[lColumn, 8] < -1.0 then
                      FItemStatsArray[lColumn, 8] := -1.0; { cap }
                  end; // if response is not FNAChar or omits...

          if (iteminfoarray[lColumn, 3] = 'P') and (itemscalearray[lColumn, 1] = 2) then
            FP := FItemStatsArray[lColumn, 2];
          if (iteminfoarray[lColumn, 3] = 'R') and (itemscalearray[lColumn, 1] = 2) then
            FP := FItemStatsArray[lColumn, 2] - 1;

          if (iteminfoarray[lColumn, 3] <> 'M') and (itemscalearray[lColumn, 1] = 2) then
          begin
            FZ := -cdfni(FP);
            FOrdinate := 0.398942 * exp(-sqr(FZ) / 2);
            FItemStatsArray[lColumn, 8] := FItemStatsArray[lColumn, 7] * (sqrt(FP * (1 - FP)) / FOrdinate);
          end;

          if FItemStatsArray[lColumn, 8] > 1.0 then
            FItemStatsArray[lColumn, 8] := 1.0; { cap }
          if FItemStatsArray[lColumn, 8] < -1.0 then
            FItemStatsArray[lColumn, 8] := -1.0; { cap }
        end; // items loop
    end; // if 2+ Domains

    for domain1 := 1 to FNumDomains do
      for domain2 := 1 to FNumDomains do
      begin
        FDomainCS := 0;
        domainmean1 := FDomainStatsArray[domain1, 2];
        domainmean2 := FDomainStatsArray[domain2, 2];

        for row := 1 to FNumberOfExaminees do
          FDomainCS := FDomainCS + (FDomainArray[row, domain1] - domainmean1) *
            (FDomainArray[row, domain2] - domainmean2);

        if FNumberOfExaminees > 1 then
          FDomainCS := FDomainCS / pred(FNumberOfExaminees);

        domainsd1 := FDomainStatsArray[domain1, 3];
        domainsd2 := FDomainStatsArray[domain2, 3];

        FDomainCorrArray[domain1, domain2] := FDomainCS / (domainsd1 * domainsd2 + 0.000000000001);
      end; // loop for domain1
  end; // for domain to FNumDomains loop Feb.2013
end;

procedure TItemanCalculation.DrawCsemGraph(nit: Integer; graph: TBitmap; title: string; csemarray: TRealArray);
var
  MaxX, mvy, incrY: Real;
  check1, j, ht, incr, GraphScaler, GraphX, GraphY, xst, yst, ntick, xadj, aj, xsp, row: Integer;
  alttext: Boolean;
  decs: string;
begin
  // Set size
  graph.LoadFromResource('graph_background', C_Pict_Width);

  with graph.Canvas do
    try
      BeginScene;

      { Set values }
      ntick := 0;
      ht := 190; // height of histogram bars!
      Stroke.Kind := TBrushKind.Solid;
      Stroke.Thickness := 1;

      Font.Family := 'Arial';
      Font.Size := 11;
      if FMaxCsem <= 2 then
        incrY := 0.25;
      if (FMaxCsem < 5) and (FMaxCsem > 2) then
        incrY := 0.5;
      if (FMaxCsem < 12) and (FMaxCsem >= 5) then
        incrY := 1;
      if (FMaxCsem < 25) and (FMaxCsem >= 12) then
        incrY := 2;
      if (FMaxCsem < 60) and (FMaxCsem >= 25) then
        incrY := 5;

      if FMaxCsem <= 2 then
        ntick := 8;
      if FMaxCsem > 2 then
        ntick := round(FMaxCsem / incrY);
      if FMaxCsem > ntick * incrY then
        ntick := ntick + 1; // so no round-down problems

      xst := 50;
      yst := 250;

      // Clear image
      // DrawBitmap(frmMain.refill.Bitmap, RectF(0, 0, 450, 300), RectF(0, 0, 450, 300), 1);

      GraphScaler := round(ht / (ntick + 1E-20)); { convert to integer }

      { Creating the appropriate Y axis labels }
      Fill.Color := TAlphaColorRec.Black;

      j := 0;
      mvy := 0;
      while mvy <= 184 do
      begin
        mvy := GraphScaler * j;
        if (incrY = 0.25) or (incrY = 0.5) then
          check1 := 9
        else
          check1 := 0;
        if (incrY = 0.25) or (incrY = 0.5) then
          decs := '0.00'
        else
          decs := '0';
        if (mvy <= 320) and (j * incrY < 0.25) then
          FillText(Rect(xst - 18 - check1, yst - round(mvy) - 7, xst - 18 - check1 + 50, yst - round(mvy) - 7 + 50),
            formatfloat(decs, 0), False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
        if (mvy <= 320) and (j * incrY < 10) and (j * incrY > 0) then
          FillText(Rect(xst - 18 - check1, yst - round(mvy) - 7, xst - 18 - check1 + 50, yst - round(mvy) - 7 + 50),
            formatfloat(decs, j * incrY), False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
        if (mvy <= 320) and (j * incrY < 100) and (j * incrY >= 10) then
          FillText(Rect(xst - 23 - check1, yst - round(mvy) - 7, xst - 23 - check1 + 50, yst - round(mvy) - 7 + 50),
            formatfloat(decs, j * incrY), False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
        if (mvy <= 320) and (j * incrY < 1000) and (j * incrY >= 100) then
          FillText(Rect(xst - 28 - check1, yst - round(mvy) - 7, xst - 28 - check1 + 50, yst - round(mvy) - 7 + 50),
            formatfloat(decs, j * incrY), False, 1, [], TTextAlign.Leading, TTextAlign.Leading);

        Stroke.Color := TAlphaColorRec.Black;
        if mvy <= 320 then
          DrawLine(Point(xst, round(yst - mvy)), Point(xst - 5, round(yst - mvy)), 1);
        j := j + 1
      end;

      { Draw axes }
      Stroke.Color := TAlphaColorRec.Black;
      DrawLine(Point(50, 250), Point(420, 250), 1); { X axis }
      DrawLine(Point(50, 250), Point(50, 37), 1); { Y axis }

      { Labels }
      Font.Family := 'Arial';
      Font.Size := 11;
      Font.Style := [];
      aj := 0;
      if nit <= 20 then
        aj := nit;
      if nit <= 20 then
        incr := 1;
      if (nit <= 40) and (nit > 20) then
        incr := 2;
      if (nit <= 60) and (nit > 40) then
        incr := 3;
      if (nit <= 100) and (nit > 60) then
        incr := 5;
      if (nit <= 200) and (nit > 100) then
        incr := 10;
      if (nit <= 400) and (nit > 200) then
        incr := 20;
      if (nit <= 500) and (nit > 400) then
        incr := 25;
      if (nit <= 1000) and (nit > 500) then
        incr := 50;
      if (nit <= 2000) and (nit > 1000) then
        incr := 100;
      if (nit <= 10000) and (nit > 2000) then
        incr := 500;
      if nit > 20 then
        aj := round(nit / incr);
      if nit > (aj * incr) then
        aj := aj + 1;
      xsp := round(350 / aj);
      for row := 0 to aj do
      begin
        xadj := 6;
        alttext := True;
        if aj > 15 then
          if 0 = (row mod 2) then
            alttext := True
          else
            alttext := False;
        if (incr * row) < 9.5 then
          xadj := 2; // adjusting the position of the xaxis labels
        if ((incr * row) >= 99.5) and ((incr * row) < 999.5) then
          xadj := 10;
        if ((incr * row) >= 999.5) and ((incr * row) < 9999.5) then
          xadj := 14;
        if alttext = True then
          if nit >= 20 then
            FillText(Rect(xst - xadj + row * xsp, yst + 8, xst - xadj + row * xsp + 50, yst + 8 + 50),
              formatfloat('0', incr * row), False, 1, [], TTextAlign.Leading, TTextAlign.Leading)
          else
            FillText(Rect(xst - xadj + row * xsp, yst + 8, xst - xadj + row * xsp + 50, yst + 8 + 50), IntToStr(incr),
              False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
        // canvas.pen.Color := RGB(22, 22, 22);
        DrawLine(Point(xst + row * xsp, yst), Point(xst + row * xsp, yst + 5), 1);
      end;
      { Actual Grouped Freq Overlay }
      GraphX := xst;
      GraphY := yst;
      MaxX := 350 * (nit / (incr * aj));
      GraphScaler := round(190 / (incrY * ntick));
      For row := 0 to nit do
      begin
        if (csemarray[row, 2] >= 0) then
        begin
          DrawLine(Point(GraphX, GraphY), Point(xst + round(row * MaxX / nit),
            yst - round(sqrt(csemarray[row, 2]) * GraphScaler)), 1);
          GraphY := yst - round(sqrt(csemarray[row, 2]) * GraphScaler);
        end
        else
          DrawLine(Point(GraphX, GraphY), Point(xst + round(row * MaxX / nit), GraphY), 1);
        GraphX := xst + round(row * (MaxX / nit));
      end;

      Font.Size := 12;
      Font.Style := [TFontStyle.fsBold];
      FillText(Rect(180, 10, 230, 60), 'CSEM', False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
      FillText(Rect(180, 275, 280, 325), 'Items', False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
      Font.Style := [];
    finally
      EndScene;
    end;
end;

procedure TItemanCalculation.DrawDistGraph(Min, Max: Real; nit, nunit, call: Integer; graph: TBitmap; title: string);
var
  // b1 : tpicture;
  MaxY, mvy, spacer, bar1: Real;
  i, j, ht, GraphScaler, GraphX, GraphY, xst, yst, MaxFr, incr, xadj, aj, xsp, row: Integer;
  alttext: Boolean;
begin
  // Set size
  graph.LoadFromResource('graph_background', C_Pict_Width);

  with graph.Canvas do
    try
      BeginScene;

      { Set values }
      ht := 190; // height of histogram bars!

      Stroke.Kind := TBrushKind.Solid;
      Stroke.Thickness := 1;

      Font.Family := 'Arial';
      Font.Size := 11;

      if (maxcategory > 2) and (call <> 4) then
      begin
        if (call = 3) and (Max - Min < 1) then
          spacer := (1 - (Max - Min)) / 2;
        if (call = 3) and (Max - Min < 1) then
          Min := Min - spacer;
        if (call = 3) and (Max - Min < 1) then
          Max := Max + spacer;
        if (call = 3) and (Max > maxcategory) then
          Max := maxcategory;
        if (IsZero = True) and (Min < 0) then
          Min := 0;
        if (IsZero = False) and (Min < 1) then
          Min := 1;
      end;
      xst := 50;
      yst := 250;

      // Clear image
      // DrawBitmap(frmMain.refill.Bitmap, RectF(0, 0, 450, 300), RectF(0, 0, 450, 300), 1);
      // DrawBitmap(frmMain.refill.Bitmap, RectF(0, 0, 380, 300), RectF(0, 0, 450, 300), 1);

      { Find the max value on Y in GroupedFreq to use for scaling the graph }
      MaxY := 0;
      for row := 1 to nbin do
      begin
        If GroupedFreqArray[row, 4] > MaxY then
          MaxY := GroupedFreqArray[row, 4];
      end;
      GraphScaler := round(ht / (MaxY + 1E-20)); { convert to integer }

      MaxFr := round(nunit * MaxY);

      incr := 1;
      if MaxFr > 12 then
      begin
        incr := 0;
        for i := 0 to 10000 do
        begin
          j := i * 10;
          if (MaxFr >= j) and (MaxFr < (j + 10)) then
          begin
            incr := i + 1;
            break;
          end;
        end;
        if (incr >= 12) and (incr < 20) then
          incr := 15;
        if (incr >= 20) and (incr < 30) then
          incr := 30;
        if (incr >= 30) and (incr < 40) then
          incr := 40;
        if (incr >= 40) and (incr < 60) then
          incr := 50;
        if (incr >= 60) and (incr < 100) then
          incr := 75;
        if (incr >= 100) and (incr < 125) then
          incr := 100;
        if (incr >= 125) and (incr < 200) then
          incr := 150;
        if (incr >= 200) and (incr < 400) then
          incr := 300;
        if (incr >= 400) and (incr < 600) then
          incr := 500;
        if (incr >= 600) and (incr < 1000) then
          incr := 800;
        if (incr >= 1000) and (incr < 2000) then
          incr := 1000;
        if (incr >= 2000) and (incr < 3000) then
          incr := 2500;
        if (incr >= 3000) and (incr < 5000) then
          incr := 4000;
        if (incr >= 5000) and (incr < 10000) then
          incr := 7500;
        if (incr >= 10000) and (incr < 50000) then
          incr := 20000;
      end;

      { Creating the appropriate Y axis labels }
      Fill.Color := TAlphaColorRec.Black;

      j := 1;
      mvy := 0;
      while mvy <= 180 do
      begin
        mvy := mvy + (GraphScaler * ((incr) / nunit));
        if (mvy <= 320) and (j * incr < 10) then
          FillText(Rect(xst - 13, yst - (round(mvy) + 7), xst - 13 + 50, yst - (round(mvy) + 7) + 50),
            formatfloat('0', j * incr), False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
        if (mvy <= 320) and (j * incr < 100) and (j * incr >= 10) then
          FillText(Rect(xst - 18, yst - (round(mvy) + 7), xst - 18 + 50, yst - (round(mvy) + 7) + 50),
            formatfloat('0', j * incr), False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
        if (mvy <= 320) and (j * incr < 1000) and (j * incr >= 100) then
          FillText(Rect(xst - 23, yst - (round(mvy) + 7), xst - 23 + 50, yst - (round(mvy) + 7) + 50),
            formatfloat('0', j * incr), False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
        if (mvy <= 320) and (j * incr < 10000) and (j * incr >= 1000) then
          FillText(Rect(xst - 28, yst - (round(mvy) + 7), xst - 28 + 50, yst - (round(mvy) + 7) + 50),
            formatfloat('0', j * incr), False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
        if (mvy <= 320) and (j * incr < 100000) and (j * incr >= 10000) then
          FillText(Rect(xst - 33, yst - (round(mvy) + 7), xst - 33 + 50, yst - (round(mvy) + 7) + 50),
            formatfloat('0', j * incr), False, 1, [], TTextAlign.Leading, TTextAlign.Leading);

        Stroke.Color := TAlphaColorRec.Black;
        if mvy <= 320 then
          DrawLine(Point(xst, round(yst - mvy)), Point(xst - 5, round(yst - mvy)), 1);
        j := j + 1
      end;

      { Draw axes }
      Stroke.Color := TAlphaColorRec.Black;
      // DrawLine(Point(50, 250), Point(360, 250), 1); { X axis }
      DrawLine(Point(50, 250), Point(420, 250), 1); { X axis }
      DrawLine(Point(50, 250), Point(50, 37), 1); { Y axis }

      { Labels }
      Font.Family := 'Arial';
      Font.Size := 11;
      Font.Style := [];
      aj := nbin;
      xsp := round(350 / aj);
      for row := 0 to aj do
      begin
        xadj := 6;
        alttext := True;
        if nbin > 15 then
          if 0 = (row mod 2) then
            alttext := True
          else
            alttext := False;
        if (Min + ((Max - Min) / 10) * row) < 9.5 then
          xadj := 2; // adjusting the position of the xaxis labels
        if ((Min + ((Max - Min) / 10) * row) >= 99.5) and ((Min + ((Max - Min) / 10) * row) < 999.5) then
          xadj := 10;
        if ((Min + ((Max - Min) / 10) * row) >= 999.5) and ((Min + ((Max - Min) / 10) * row) < 9999.5) then
          xadj := 14;
        if (nit >= 20) and (brange < 20) then
          i := 0
        else
          i := 1;
        if (alttext = True) and (call = 1) then
          if brange >= 20 then
            FillText(Rect(xst - xadj + row * xsp, yst + 8, xst - xadj + row * xsp + 50, yst + 8 + 50),
              formatfloat('0', (Min + ((Max - Min) / (aj + 1E-20)) * row)), False, 1, [], TTextAlign.Leading,
              TTextAlign.Leading)
          else
            FillText(Rect(xst - xadj + row * xsp, yst + 8, xst - xadj + row * xsp + 50, yst + 8 + 50),
              IntToStr(round(Max + i) - (aj - row)), False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
        // pen.Color := RGB(22, 22, 22);
        DrawLine(PointF(xst + row * xsp, yst), PointF(xst + row * xsp, yst + 5), 1);
      end;
      if call = 2 then
        for j := 0 to nbin do
          FillText(Rect(xst - 6 + j * xsp, yst + 8, xst - 6 + j * xsp + 50, yst + 8 + 50),
            formatfloat('0.0', 0 + 0.1 * j), False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
      if call = 3 then
        for j := 0 to nbin do
          FillText(Rect(xst - 6 + j * xsp, yst + 8, xst - 6 + j * xsp + 50, yst + 8 + 50),
            formatfloat('0.0', Min + (brange / (aj + 1E-20)) * j), False, 1, [], TTextAlign.Leading,
            TTextAlign.Leading);
      if call = 4 then
        for j := 0 to nbin do
        begin
          if Min + 0.1 * j < 0 then
            aj := -3
          else
            aj := 0;
          bar1 := 0.1;
          if nbin >= 13 then
            if j mod 2 = 0 then
              FillText(RectF(xst - 6 + aj + j * xsp, yst + 8, xst - 6 + aj + j * xsp + 50, yst + 8 + 50),
                formatfloat('0.0', Min + bar1 * j), False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
          if nbin <= 12 then
            FillText(RectF(xst - 6 + aj + j * xsp, yst + 8, xst - 6 + aj + j * xsp + 50, yst + 8 + 50),
              formatfloat('0.0', Min + bar1 * j), False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
        end;

      { Actual Grouped Freq Overlay }
      GraphX := xst;
      GraphY := yst;
      For row := 1 to nbin do
      begin
        DrawLine(Point(GraphX, GraphY), Point(GraphX, yst - round(GroupedFreqArray[row, 4] * GraphScaler)), 1);
        GraphY := yst - round(GroupedFreqArray[row, 4] * GraphScaler);
        DrawLine(Point(GraphX, GraphY), Point(GraphX + xsp, GraphY), 1);
        GraphX := GraphX + xsp;
        // xsp needs to be adjusted if nbin equals 20 or 18
        DrawLine(Point(GraphX, GraphY), Point(GraphX, yst), 1);
      end;

      Font.Size := 12;
      Font.Style := [TFontStyle.fsBold];
      FillText(Rect(6, 140, 56, 190), 'N', False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
      FillText(Rect(180, 275, 280, 325), title, False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
      Font.Style := [];
    finally
      EndScene;
    end;
end;

procedure TItemanCalculation.DrawItemGraph(graph: TBitmap; item: Integer; itemscalearray: TIntegerArray;
  itemidarray: TStringArray; key1, start0: Char);
var
  GraphX, GraphY, tmpX, tmpY: Integer;
  altcap, i, j, xspan, max1, mf, z: Integer;
  storage, store2: array [1 .. 9] of Real;
  altno, ltr: array [1 .. 9] of Integer;
  colo, ideq: string;
begin
  // Set size
  graph.LoadFromResource('graph_background', C_Pict_Width);

  with graph.Canvas do
    try
      BeginScene;

      { Set values }
      Stroke.Kind := TBrushKind.Solid;
      Stroke.Thickness := 2;
      Font.Family := 'Arial';
      Font.Size := 12;
      Fill.Color := TAlphaColorRec.Black;

      // Clear image
      // DrawBitmap(frmMain.refill.Bitmap, RectF(0, 0, 450, 300), RectF(0, 0, 450, 300), 1);

      altcap := itemscalearray[item, 1];
      if altcap > 9 then
        altcap := 9;

      for i := 1 to altcap do
        store2[i] := FSubGroupPArray[i * FNumCut];

      for j := 1 to altcap do
      begin
        for i := 1 to altcap do
        begin
          if i = 1 then
            storage[j] := store2[i];
          if i = 1 then
            max1 := 1;
          if store2[i] > storage[j] then
          begin
            storage[j] := store2[i];
            max1 := i;
          end;
        end;
        store2[max1] := -1;
        altno[j] := max1;
      end;

      for j := 1 to altcap do
        ltr[j] := 0;

      for i := 1 to altcap - 1 do
      begin
        for j := i to altcap - 1 do
        begin
          if storage[i] - storage[j + 1] < 0.07 then
            ltr[altno[j + 1]] := ltr[altno[j + 1]] + 9
          else
            ltr[altno[j + 1]] := ltr[altno[j + 1]] + 0;
        end;
        if storage[1] - storage[i + 1] >= 0.07 then
          ltr[altno[i + 1]] := 0;
      end;

      mf := altcap;
      for j := 1 to 5 do
      begin

        for i := 1 to mf - 1 do
          if (storage[i] - storage[i + 1] < 0.07) and (ltr[altno[i]] = ltr[altno[i + 1]]) then
            ltr[altno[i + 1]] := ltr[altno[i + 1]] + 9;
        for i := 1 to mf - 2 do
          if (storage[i] - storage[i + 2] < 0.07) and (ltr[altno[i]] = ltr[altno[i + 2]]) then
            ltr[altno[i + 2]] := ltr[altno[i + 2]] + 9;
        for i := 1 to mf - 3 do
          if mf > 3 then
            if (storage[i] - storage[i + 3] < 0.07) and (ltr[altno[i]] = ltr[altno[i + 3]]) then
              ltr[altno[i + 3]] := ltr[altno[i + 3]] + 9;
        for i := 1 to mf - 4 do
          if mf > 4 then
            if (storage[i] - storage[i + 4] < 0.07) and (ltr[altno[i]] = ltr[altno[i + 4]]) then
              ltr[altno[i + 4]] := ltr[altno[i + 4]] + 9;
        for i := 1 to mf - 5 do
          if mf > 5 then
            if (storage[i] - storage[i + 5] < 0.07) and (ltr[altno[i]] = ltr[altno[i + 5]]) then
              ltr[altno[i + 5]] := ltr[altno[i + 5]] + 9;
        for i := 1 to mf - 6 do
          if mf > 6 then
            if (storage[i] - storage[i + 6] < 0.07) and (ltr[altno[i]] = ltr[altno[i + 6]]) then
              ltr[altno[i + 6]] := ltr[altno[i + 6]] + 9;
        for i := 1 to mf - 7 do
          if mf > 7 then
            if (storage[i] - storage[i + 7] < 0.07) and (ltr[altno[i]] = ltr[altno[i + 7]]) then
              ltr[altno[i + 7]] := ltr[altno[i + 7]] + 9;
        for i := 1 to mf - 8 do
          if mf > 8 then
            if (storage[i] - storage[i + 8] < 0.07) and (ltr[altno[i]] = ltr[altno[i + 8]]) then
              ltr[altno[i + 8]] := ltr[altno[i + 8]] + 9;

      end;
      { Draw quantile lines: line 250 is 0.0, line 50 is 1.0, so take P*200 }
      For i := 1 to altcap do
      begin
        if key1 = '-' then
          z := (itemscalearray[item, 1] + 1) - i
        else
          z := i;
        if z = 1 then
          Stroke.Color := TAlphaColorRec.Alpha or (((((215 shl 8) or 95) shl 8) or 68) shl 8);
        if z = 2 then
          Stroke.Color := TAlphaColorRec.Navy; // TAlphaColorRec.Green;
        if z = 3 then
          Stroke.Color := TAlphaColorRec.Gold; // TAlphaColorRec.Blue;
        if z = 4 then
          Stroke.Color := TAlphaColorRec.Gray; // TAlphaColorRec.Olive;
        if z = 5 then
          Stroke.Color := TAlphaColorRec.Lightblue; // TAlphaColorRec.Dkgray;
        if z = 6 then
          Stroke.Color := TAlphaColorRec.Navy;
        if z = 7 then
          Stroke.Color := TAlphaColorRec.Red;
        if z = 8 then
          Stroke.Color := TAlphaColorRec.Teal;
        if z = 9 then
          Stroke.Color := TAlphaColorRec.Purple;

        xspan := round(250 / (FNumCut - 1));
        tmpY := round(200 * FSubGroupPArray[((i - 1) * FNumCut + 1)]);
        GraphX := 100;
        GraphY := 250 - tmpY;
        DrawEllipse(Rect(98, GraphY - 2, 102, GraphY + 2), 1);
        for j := 2 to FNumCut do
        begin
          tmpY := round(200 * FSubGroupPArray[((i - 1) * FNumCut + j)]);
          DrawEllipse(Rect((j - 1) * xspan + 98, 250 - tmpY - 2, (j - 1) * xspan + 102, 250 - tmpY + 2), 1);
          DrawLine(Point(GraphX, GraphY), Point(xspan * (j - 1) + 100, 250 - tmpY), 1);
          GraphX := xspan * (j - 1) + 100;
          GraphY := 250 - tmpY;
        end;
        GraphY := tmpY;

        { Label the lines with the response }
        if GraphY < 6 then
          GraphY := 6;

        if start0 = 'P' then
        begin
          if z = 1 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), '0', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 2 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), '1', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 3 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), '2', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 4 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), '3', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 5 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), '4', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 6 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), '5', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 7 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), '6', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 8 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), '7', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 9 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), '8', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
        end;
        if (FRespAsInt = True) and (start0 <> 'P') then
        begin
          if z = 1 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), '1', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 2 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), '2', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 3 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), '3', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 4 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), '4', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 5 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), '5', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 6 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), '6', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 7 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), '7', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 8 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), '8', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 9 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), '9', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
        end;
        if (FRespAsInt = False) and (start0 <> 'P') then
        begin
          if z = 1 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), 'A', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 2 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), 'B', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 3 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), 'C', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 4 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), 'D', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 5 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), 'E', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 6 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), 'F', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 7 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), 'G', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 8 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), 'H', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
          if z = 9 then
            FillText(Rect(314 + ltr[i], 242 - GraphY, 314 + ltr[i] + 50, 242 - GraphY + 50), 'I', False, 1, [],
              TTextAlign.Leading, TTextAlign.Leading);
        end;
      end;

      { Draw axes }
      Stroke.Color := TAlphaColorRec.Alpha or (((((22 shl 8) or 22) shl 8) or 22) shl 8);
      DrawLine(Point(50, 250), Point(420, 250), 1); { X axis }
      DrawLine(Point(50, 250), Point(50, 50), 1); { Y axis }

      GraphX := 100;
      DrawLine(Point(GraphX, 245), Point(100, 255), 1);
      { Hash marks... }

      for j := 2 to FNumCut do
      begin
        GraphX := 100 + (j - 1) * xspan;
        DrawLine(Point(GraphX, 245), Point(GraphX, 255), 1);
        { Hash marks... }
      end;

      { Labels }
      Font.Family := 'Arial';
      Font.Size := 14;
      // canvas.brush.color := RGB(250, 250, 198);
      FillText(Rect(20, 40, 70, 90), '1.0', False, 1, [], TTextAlign.Leading, TTextAlign.Leading); { max Y }
      // canvas.brush.color := RGB(250, 250, 198);
      FillText(Rect(20, 140, 70, 190), '0.5', False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
      // canvas.brush.color := RGB(250, 250, 198);
      FillText(Rect(20, 240, 70, 290), '0.0', False, 1, [], TTextAlign.Leading, TTextAlign.Leading); { min Y }
      for j := 1 to FNumCut do
        FillText(Rect(95 + (j - 1) * xspan, 260, 95 + (j - 1) * xspan + 50, 310), IntToStr(j), False, 1, [],
          TTextAlign.Leading, TTextAlign.Leading);
      Font.Size := 14;
      FillText(Rect(177, 275, 227, 325), 'Group', False, 1, [], TTextAlign.Leading, TTextAlign.Leading); { X name }
      // canvas.brush.color := RGB(250, 250, 198);
      FillText(Rect(10, 120, 60, 170), 'P', False, 1, [], TTextAlign.Leading, TTextAlign.Leading); { Y name }
      // canvas.brush.color := RGB(250, 250, 198);
      if IntToStr(item) = itemidarray[item] then
        ideq := ''
      else
        ideq := itemidarray[item];
      if IntToStr(item) = itemidarray[item] then
        colo := ''
      else
        colo := ': ';
      if IntToStr(item) = itemidarray[item] then
        mf := 25
      else
        mf := 0;
      FillText(Rect(150 + mf, 10, 250 + mf, 60), 'Item ' + IntToStr(item) + colo + ideq, False, 1, [],
        TTextAlign.Leading, TTextAlign.Leading); { Title }

      Font.Size := 11;
      FillText(Rect(380, 282, 445, 332), '© 2021 ASC', False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
      // 169 is copyright   chr(0169)       ©
      // Picture.SaveToFile('Sample Files\'+IntToStr(GetTickCount)+'.jpg');
    finally
      EndScene;
    end; // with graph do begin
end;

procedure TItemanCalculation.DrawScatterplot(Min, Max: Real; nit, nunit, call: Integer; graph: TBitmap;
  iteminfoarray: TChar2Array; type1: Char; title: string);
var
  spacer, BndR, minRc, abscorr: Real;
  j, ht, GraphScaler, GraphX, GraphY, xst, yst, aj, xsp, ii, row: Integer;
begin
  // Set size
  graph.LoadFromResource('graph_background', C_Pict_Width);

  with graph.Canvas do
    try
      BeginScene;
      { Set values }
      ht := 200; // height of histogram bars!
      Stroke.Kind := TBrushKind.Solid;
      Stroke.Thickness := 1;

      Font.Family := 'Arial';
      Font.Size := 11;

      if maxcategory > 2 then
      begin
        if (call = 2) and (Max - Min < 1) then
          spacer := (1 - (Max - Min)) / 2;
        if (call = 2) and (Max - Min < 1) then
          Min := Min - spacer;
        if (call = 2) and (Max - Min < 1) then
          Max := Max + spacer;
        if Max > maxcategory then
          Max := maxcategory;
        if (IsZero = True) and (Min < 0) then
          Min := 0;
        if (IsZero = False) and (Min < 1) then
          Min := 1;
      end;
      xst := 50;
      yst := 250;

      // Clear image
      // DrawBitmap(frmMain.refill.Bitmap, RectF(0, 0, 450, 300), RectF(0, 0, 450, 300), 1);

      if type1 = 'M' then
        minRc := FMinRPBis
      else
        minRc := FMinCorr;

      BndR := 1;
      { Find the max value on Y in GroupedFreq to use for scaling the graph }
      if (minRc < 0.05) and (minRc > -0.05) then
        BndR := 1 + 0.1;
      if (minRc < -0.05) and (minRc > -0.15) then
        BndR := 1 + 0.2;
      if (minRc < -0.15) and (minRc > -0.25) then
        BndR := 1 + 0.3;
      if (minRc < -0.25) and (minRc > -0.35) then
        BndR := 1 + 0.4;
      if (minRc < -0.35) and (minRc > -0.45) then
        BndR := 1 + 0.5;
      if (minRc < -0.45) and (minRc > -0.55) then
        BndR := 1 + 0.6;
      if (minRc < -0.55) and (minRc > -0.65) then
        BndR := 1 + 0.7;
      if (minRc < -0.65) and (minRc > -0.75) then
        BndR := 1 + 0.8;
      if (minRc < -0.75) and (minRc > -0.85) then
        BndR := 1 + 0.9;
      if (minRc < -0.85) and (minRc > -0.95) then
        BndR := 1 + 1;
      GraphScaler := round(ht / BndR); { convert to integer }

      { Draw axes }
      Stroke.Color := TAlphaColorRec.Black;
      DrawLine(Point(50, 250), Point(420, 250), 1); { X axis }
      DrawLine(Point(50, 250), Point(50, 37), 1); { Y axis }

      { Labels }
      Font.Family := 'Arial';
      Font.Size := 11;
      Font.Style := [];
      Fill.Color := TAlphaColorRec.Black;

      aj := nbin;
      xsp := round(350 / aj);
      for row := 0 to aj do
      begin
        // pen.Color := RGB(22, 22, 22);
        DrawLine(Point(xst + row * xsp, yst), Point(xst + row * xsp, yst + 5), 1);
      end;
      if call = 2 then
        brange := round(Max - Min);
      if call = 1 then
        for j := 0 to nbin do
          FillText(Rect((xst - 6) + j * xsp, (yst + 8), (xst - 6) + j * xsp + 50, (yst + 8) + 50),
            formatfloat('0.0', 0 + 0.1 * j), False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
      if call = 2 then
        for j := 0 to nbin do
          FillText(Rect((xst - 6) + j * xsp, (yst + 8), (xst - 6) + j * xsp + 50, (yst + 8) + 50),
            formatfloat('0.0', Min + (brange / aj) * j), False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
      for j := 1 to round(BndR * 10) do
      begin
        if (j * 0.1) - (BndR - 1) < -0.01 then
          ii := -3
        else
          ii := 0;
        if (BndR >= 1.3) and (j mod 2 <> 0) then
          FillText(Rect(30 + ii, round(yst - (j * 0.1) * GraphScaler) - 8, 30 + ii + 50,
            round(yst - (j * 0.1) * GraphScaler) - 8 + 50), formatfloat('0.0', (j * 0.1) - (BndR - 1)), False, 1, [],
            TTextAlign.Leading, TTextAlign.Leading);
        if BndR < 1.3 then
          FillText(Rect(30 + ii, round(yst - (j * 0.1) * GraphScaler) - 8, 30 + ii + 50,
            round(yst - (j * 0.1) * GraphScaler) - 8 + 50), formatfloat('0.0', (j * 0.1) - (BndR - 1)), False, 1, [],
            TTextAlign.Leading, TTextAlign.Leading);
        DrawLine(Point(50, round(yst - (j * 0.1) * GraphScaler)), Point(46, round(yst - (j * 0.1) * GraphScaler)), 1);
      end;
      { Actual Grouped Freq Overlay }
      GraphX := xst;
      GraphY := yst;
      Stroke.Color := TAlphaColorRec.Black;
      For row := 1 to FNumberOfItems do
        if iteminfoarray[row, 2] = 'Y' then
        begin
          frmMain.ProgressBar.Value := frmMain.ProgressBar.Value + 1;
          Application.ProcessMessages;
          // Canvas.MoveTo(xst, yst);  //???
          abscorr := FItemStatsArray[row, 3];
          if abscorr > 0 then
            abscorr := abscorr + (BndR - 1); // if no near 0 r, BndR - 1 = 0
          if abscorr <= 0 then
            abscorr := abscorr + (BndR - 1);
          GraphY := yst - round(abscorr * GraphScaler);
          if (type1 = 'M') and (iteminfoarray[row, 3] = 'M') then
            GraphX := xst + round(FItemStatsArray[row, 2] * 300);
          if (type1 <> 'M') and (iteminfoarray[row, 3] <> 'M') then
            GraphX := xst + round(FItemStatsArray[row, 2] * (300 / Max));
          if (type1 = 'M') and (iteminfoarray[row, 3] = 'M') then
            DrawEllipse(Rect(GraphX - 2, GraphY - 2, GraphX + 2, GraphY + 2), 1);
          if (type1 <> 'M') and (iteminfoarray[row, 3] <> 'M') then
            DrawEllipse(Rect(GraphX - 2, GraphY - 2, GraphX + 2, GraphY + 2), 1);
        end;

      Font.Size := 12;
      Font.Style := [TFontStyle.fsBold];
      if type1 = 'M' then
        FillText(Rect(2, 140, 52, 190), 'Rpb', False, 1, [], TTextAlign.Leading, TTextAlign.Leading)
      else
        FillText(Rect(6, 140, 56, 190), 'R', False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
      FillText(Rect(180, 275, 280, 325), title, False, 1, [], TTextAlign.Leading, TTextAlign.Leading);
      Font.Style := [];
    finally
      EndScene;
    end;
end;

function TItemanCalculation.EstimateKR20(Domain: Integer; iteminfoarray: TChar2Array;
  itemscalearray: TIntegerArray): Real;
var
  lRow: Integer;
begin
  FPSum := 0.0;

  for lRow := 1 to FNumberOfItems do
    if (iteminfoarray[lRow, 2] = 'Y') and (itemscalearray[lRow, 2] = Domain) then
    begin
      FPTemp := FItemStatsArray[lRow, 2];
      FPSum := FPSum + FItemStatsArray[lRow, 6];
      // FItemStatsArray[lRow,6] has item variance
    end;

  if FDomainStatsArray[Domain, 1] > 1 then
    Result := (FDomainStatsArray[Domain, 1] / (FDomainStatsArray[Domain, 1] - 1)) *
      (1 - (FPSum / ((FDomainStatsArray[Domain, 3] * FDomainStatsArray[Domain, 3]) + 0.000000000001)))
  else
    Result := 0;
end;

function TItemanCalculation.CalculateMeanP(Domain: Integer; iteminfoarray: TChar2Array;
  itemscalearray: TIntegerArray): Real;
var
  lRow: Integer;
begin
  FPSum := 0.0;

  for lRow := 1 to FNumberOfItems do
  begin
    If (iteminfoarray[lRow, 2] = 'Y') and (iteminfoarray[lRow, 3] = 'M') then
      If itemscalearray[lRow, 2] = Domain then
      begin
        FPTemp := FItemStatsArray[lRow, 2];
        FPSum := FPSum + FPTemp;
      end;
  end;
  Result := FPSum / (FDomMult[Domain] + 0.000000000001);
end;

procedure TItemanCalculation.CalcGF(Min, Max: Real; call: Integer; scorearray: TIntegerArray);
var
  binc, row, counter, rawsc: Integer;
  TableInterval: Real;
  { Calculate Grouped Frequency Table }
begin
  for counter := 1 to nbin do
  begin
    for row := 1 to 4 do
      GroupedFreqArray[counter, row] := 0;
  end; // clearing the array

  binc := nbin - 1;
  if (Max - Min < 20) and (Max < 20) then
    binc := binc - 1;
  if (Max - Min < 20) and (Max < 20) then
    TableInterval := 1
  else
    TableInterval := (Max - Min) / binc;
  GroupedFreqArray[1, 1] := Min - 0.00000000001;
  GroupedFreqArray[nbin, 2] := Max + 0.0000000001;
  GroupedFreqArray[1, 3] := 0;
  For counter := 2 to nbin do
  begin
    GroupedFreqArray[counter, 1] := Min + (counter - 1) * TableInterval;
    GroupedFreqArray[counter - 1, 2] := GroupedFreqArray[counter, 1] - 0.00000000000001;
    GroupedFreqArray[counter, 3] := 0;
  end;
  frmMain.ProgressBar.Value := frmMain.ProgressBar.Value + 1;
  Application.ProcessMessages;

  For counter := 1 to FNumberOfExaminees do
  begin
    For row := 1 to nbin do
    begin
      if call = 1 then
        rawsc := scorearray[counter, 1]
      else
        rawsc := FDomainArray[counter, call - 1];
      If rawsc >= GroupedFreqArray[row, 1] then
      begin
        If rawsc < GroupedFreqArray[row, 2] then
        begin
          GroupedFreqArray[row, 3] := GroupedFreqArray[row, 3] + 1;
        end;
      end;
    end;
  end;
  For counter := 1 to nbin do
  begin
    GroupedFreqArray[counter, 4] := GroupedFreqArray[counter, 3] / FNumberOfExaminees;
  end;
end;

procedure TItemanCalculation.CalcItemGF(Min, Max: Real; call, denom: Integer; type1: Char; iteminfoarray: TChar2Array);
var
  row, counter: Integer;
  TableInterval, rawsc, spacer: Real;
  { Calculate Grouped Frequency Table }
begin
  for counter := 1 to nbin do
  begin
    for row := 1 to 4 do
      GroupedFreqArray[counter, row] := 0;
  end; // clearing the array

  if (type1 <> 'M') and (Max - Min < 1) then
    TableInterval := 0.1
  else
    TableInterval := (Max - Min) / (nbin + 1E-20);
  if (maxcategory > 2) and (call <> 4) then
  begin
    if (type1 <> 'M') and (Max - Min < 1) then
      spacer := (1 - (Max - Min)) / 2;
    if (type1 <> 'M') and (Max - Min < 1) then
      Min := Min - spacer;
    if (type1 <> 'M') and (Max - Min < 1) then
      Max := Max + spacer;
  end;
  GroupedFreqArray[1, 1] := Min - 0.00000000001;
  GroupedFreqArray[nbin, 2] := Max + 0.0000000001;
  GroupedFreqArray[1, 3] := 0;
  For counter := 2 to nbin do
  begin
    GroupedFreqArray[counter, 1] := Min + (counter - 1) * TableInterval;
    GroupedFreqArray[counter - 1, 2] := GroupedFreqArray[counter, 1] - 0.0000000000000000000000000001;
    GroupedFreqArray[counter, 3] := 0;
  end;

  For counter := 1 to FNumberOfItems do
  begin
    if (iteminfoarray[counter, 3] = 'P') and (type1 = 'R') then
      type1 := 'P';
    if (iteminfoarray[counter, 3] = 'R') and (type1 = 'P') then
      type1 := 'R';
    if (iteminfoarray[counter, 3] = type1) and (iteminfoarray[counter, 2] = 'Y') then
    begin // type1 is set to be M or R
      For row := 1 to nbin do
      begin
        if call = 1 then
          rawsc := FItemStatsArray[counter, 2]
        else
          rawsc := FItemStatsArray[counter, 3];
        // writeln(outputfile, rawsc:20:23, ',', GroupedFreqArray[row,1]:20:23, ',', GroupedFreqArray[row,2]:20:23, ',', GroupedFreqArray[row,3]:20:23);     //QA
        If rawsc - GroupedFreqArray[row, 2] < 0.0000000001 then
          rawsc := rawsc + 0.0000000001;
        // this helps when 0.7 is stored as 0.699999999999999 as bytes
        if rawsc > 1.0 then
          rawsc := 1.0;
        If rawsc >= GroupedFreqArray[row, 1] then
        begin
          If rawsc < GroupedFreqArray[row, 2] then
          begin
            GroupedFreqArray[row, 3] := GroupedFreqArray[row, 3] + 1;
          end;
        end;
      end;
    end;
  end;
  For counter := 1 to nbin do
  begin
    GroupedFreqArray[counter, 4] := GroupedFreqArray[counter, 3] / (denom + 1E-20);
  end;
end;

function TItemanCalculation.CalculateItemMean(Domain: Integer; iteminfoarray: TChar2Array;
  itemscalearray: TIntegerArray): Real;
var
  lRow: Integer;
begin
  FPSum := 0.0;

  for lRow := 1 to FNumberOfItems do
    if (iteminfoarray[lRow, 2] = 'Y') and (iteminfoarray[lRow, 3] <> 'M') and (itemscalearray[lRow, 2] = Domain) then
    begin
      FPTemp := FItemStatsArray[lRow, 2];
      FPSum := FPSum + FPTemp;
    end;

  Result := FPSum / (FDomRate[Domain] + 0.000000000001);
end;

function TItemanCalculation.CalculateMeanRpbis(Domain: Integer; iteminfoarray: TChar2Array;
  itemscalearray: TIntegerArray): Real;
var
  lRow: Integer;
begin
  FPSum := 0.0;

  for lRow := 1 to FNumberOfItems do
    if (iteminfoarray[lRow, 2] = 'Y') and (itemscalearray[lRow, 2] = Domain) then
    begin
      FPTemp := FItemStatsArray[lRow, 3];
      FPSum := FPSum + FPTemp;
    end;

  Result := FPSum / (FDomainStatsArray[Domain, 1] + 0.000000000001);
end;

procedure TItemanCalculation.Reliability(splitscore1, splitscore2, splitdomain1, splitdomain2,
  itemscalearray: TIntegerArray; iteminfoarray, responsearray: TChar2Array);
var
  row, column, i, j, ticker, v, cnum: Integer;
  num1, num2, rand: Real;
  sd1, sd2, CP1, CP2, CP3, CP4, CP5, CP6: Real;
  temp, tempM, sumsq: array [1 .. 12] of Real;
  domsum, domM, domsq: array [1 .. 6, 1 .. 50] of Real;
  domcp: array [1 .. 3, 1 .. 50] of Real;
  domnum: array [1 .. 50] of Real;
  lDomain: Integer;
begin
  Randomize;

  FIX := random(30000);
  FIY := random(30000);
  FIZ := random(30000);
  i := 0;
  j := 0;

  for row := 1 to 12 do
  begin
    sumsq[row] := 0;
    temp[row] := 0;
    tempM[row] := 0;
  end;

  CP1 := 0;
  CP2 := 0;
  CP3 := 0;
  CP4 := 0;
  CP5 := 0;
  CP6 := 0;

  for row := 1 to 6 do
    for column := 1 to 50 do
    begin
      domsum[row, column] := 0;
      domsq[row, column] := 0;
      domM[row, column] := 0;
      domnum[column] := 0;
    end;

  for row := 1 to 3 do
    for column := 1 to 50 do
      domcp[row, column] := 0;

  num1 := FValidItems / 2;
  num2 := FScoredItems / 2;
  ticker := 0;

  for column := 1 to FNumberOfItems do
    if iteminfoarray[column, 2] <> 'N' then
    begin
      rand := RandomNum(column);

      if i >= num1 then
        rand := 0; // force item into FGroup 2 if half in FGroup 1
      if j >= num1 then
        rand := 1; // force item into FGroup 1 if half in FGroup 2
      if rand >= 0.50 then
        i := i + 1;
      if rand < 0.50 then
        j := j + 1;

      ticker := ticker + 1;
      frmMain.ProgressBar.Value := frmMain.ProgressBar.Value + 1;
      Application.ProcessMessages;

      for row := 1 to FNumberOfExaminees do
        if (UpCase(responsearray[row, column]) <> UpCase(FNAChar)) and
          (UpCase(responsearray[row, column]) <> UpCase(FOmitChar)) then // omits and nas are excluded
        begin
          FCh := responsearray[row, column];

          if rand >= 0.50 then
          begin
            if iteminfoarray[column, 3] = 'M' then
              if UpCase(iteminfoarray[column, 1]) = UpCase(responsearray[row, column]) then
                splitscore1[row, 1] := splitscore1[row, 1] + 1;

            if iteminfoarray[column, 3] <> 'M' then
              splitscore1[row, 1] := splitscore1[row, 1] + ConvertResponseToInteger(FCh, iteminfoarray, column,
                itemscalearray[column, 1]);
          end;

          if rand < 0.50 then
          begin
            if iteminfoarray[column, 3] = 'M' then
              if UpCase(iteminfoarray[column, 1]) = UpCase(responsearray[row, column]) then
                splitscore2[row, 1] := splitscore2[row, 1] + 1;

            if iteminfoarray[column, 3] <> 'M' then
              splitscore2[row, 1] := splitscore2[row, 1] + ConvertResponseToInteger(FCh, iteminfoarray, column,
                itemscalearray[column, 1]);
          end;

          if ticker <= num1 then
          begin
            if iteminfoarray[column, 3] = 'M' then
              if UpCase(iteminfoarray[column, 1]) = UpCase(responsearray[row, column]) then
                splitscore1[row, 3] := splitscore1[row, 3] + 1;

            if iteminfoarray[column, 3] <> 'M' then
              splitscore1[row, 3] := splitscore1[row, 3] + ConvertResponseToInteger(FCh, iteminfoarray, column,
                itemscalearray[column, 1]);
          end;

          if ticker > num1 then
          begin
            if iteminfoarray[column, 3] = 'M' then
              if UpCase(iteminfoarray[column, 1]) = UpCase(responsearray[row, column]) then
                splitscore2[row, 3] := splitscore2[row, 3] + 1;

            if iteminfoarray[column, 3] <> 'M' then
              splitscore2[row, 3] := splitscore2[row, 3] + ConvertResponseToInteger(FCh, iteminfoarray, column,
                itemscalearray[column, 1]);
          end;

          if (ticker mod 2) > 0 then
          begin
            if iteminfoarray[column, 3] = 'M' then
              if UpCase(iteminfoarray[column, 1]) = UpCase(responsearray[row, column]) then
                splitscore1[row, 5] := splitscore1[row, 5] + 1;

            if iteminfoarray[column, 3] <> 'M' then
              splitscore1[row, 5] := splitscore1[row, 5] + ConvertResponseToInteger(FCh, iteminfoarray, column,
                itemscalearray[column, 1]);
          end;

          if (ticker mod 2) = 0 then
          begin
            if iteminfoarray[column, 3] = 'M' then
              if UpCase(iteminfoarray[column, 1]) = UpCase(responsearray[row, column]) then
                splitscore2[row, 5] := splitscore2[row, 5] + 1;

            if iteminfoarray[column, 3] <> 'M' then
              splitscore2[row, 5] := splitscore2[row, 5] + ConvertResponseToInteger(FCh, iteminfoarray, column,
                itemscalearray[column, 1]);
          end;

        end; // for row loop
    end; // for column loop -- items loop

  if FPreTest >= 1 then
  begin
    i := 0;
    j := 0;
    ticker := 0;

    for column := 1 to FNumberOfItems do
      if iteminfoarray[column, 2] = 'Y' then
      begin
        rand := RandomNum(column);

        if i >= num2 then
          rand := 0; // force item into FGroup 2 if half in FGroup 1
        if j >= num2 then
          rand := 1; // force item into FGroup 1 if half in FGroup 2
        if rand >= 0.50 then
          i := i + 1;
        if rand < 0.50 then
          j := j + 1;

        ticker := ticker + 1;
        frmMain.ProgressBar.Value := frmMain.ProgressBar.Value + 1;
        Application.ProcessMessages;

        for row := 1 to FNumberOfExaminees do
          if (UpCase(responsearray[row, column]) <> UpCase(FNAChar)) and
            (UpCase(responsearray[row, column]) <> UpCase(FOmitChar)) then
          begin // omits and nas are excluded
            FCh := responsearray[row, column];

            if rand >= 0.50 then
            begin
              if iteminfoarray[column, 3] = 'M' then
                if UpCase(iteminfoarray[column, 1]) = UpCase(responsearray[row, column]) then
                  splitscore1[row, 2] := splitscore1[row, 2] + 1;

              if iteminfoarray[column, 3] <> 'M' then
                splitscore1[row, 2] := splitscore1[row, 2] + ConvertResponseToInteger(FCh, iteminfoarray, column,
                  itemscalearray[column, 1]);
            end;

            if rand < 0.50 then
            begin
              if iteminfoarray[column, 3] = 'M' then
                if UpCase(iteminfoarray[column, 1]) = UpCase(responsearray[row, column]) then
                  splitscore2[row, 2] := splitscore2[row, 2] + 1;

              if iteminfoarray[column, 3] <> 'M' then
                splitscore2[row, 2] := splitscore2[row, 2] + ConvertResponseToInteger(FCh, iteminfoarray, column,
                  itemscalearray[column, 1]);
            end;

            if ticker <= num2 then
            begin
              if iteminfoarray[column, 3] = 'M' then
                if UpCase(iteminfoarray[column, 1]) = UpCase(responsearray[row, column]) then
                  splitscore1[row, 4] := splitscore1[row, 4] + 1;

              if iteminfoarray[column, 3] <> 'M' then
                splitscore1[row, 4] := splitscore1[row, 4] + ConvertResponseToInteger(FCh, iteminfoarray, column,
                  itemscalearray[column, 1]);
            end;

            if ticker > num2 then
            begin
              if iteminfoarray[column, 3] = 'M' then
                if UpCase(iteminfoarray[column, 1]) = UpCase(responsearray[row, column]) then
                  splitscore2[row, 4] := splitscore2[row, 4] + 1;

              if iteminfoarray[column, 3] <> 'M' then
                splitscore2[row, 4] := splitscore2[row, 4] + ConvertResponseToInteger(FCh, iteminfoarray, column,
                  itemscalearray[column, 1]);
            end;

            if (ticker mod 2) > 0 then
            begin
              if iteminfoarray[column, 3] = 'M' then
                if UpCase(iteminfoarray[column, 1]) = UpCase(responsearray[row, column]) then
                  splitscore1[row, 6] := splitscore1[row, 6] + 1;

              if iteminfoarray[column, 3] <> 'M' then
                splitscore1[row, 6] := splitscore1[row, 6] + ConvertResponseToInteger(FCh, iteminfoarray, column,
                  itemscalearray[column, 1]);
            end;

            if (ticker mod 2) = 0 then
            begin
              if iteminfoarray[column, 3] = 'M' then
                if UpCase(iteminfoarray[column, 1]) = UpCase(responsearray[row, column]) then
                  splitscore2[row, 6] := splitscore2[row, 6] + 1;

              if iteminfoarray[column, 3] <> 'M' then
                splitscore2[row, 6] := splitscore2[row, 6] + ConvertResponseToInteger(FCh, iteminfoarray, column,
                  itemscalearray[column, 1]);
            end;
          end; // for row loop
      end; // for column loop -- items loop
  end; // if FPreTest is >= 1 (there are FPreTest items!

  if FNumDomains > 1 then
    for lDomain := 1 to FNumDomains do
    begin
      i := 0;
      j := 0;
      ticker := 0;
      domnum[lDomain] := FDomainStatsArray[lDomain, 1] / 2;
      cnum := 3 * pred(lDomain);

      for column := 1 to FNumberOfItems do
        if (itemscalearray[column, 2] = lDomain) and (iteminfoarray[column, 2] = 'Y') then
        begin
          rand := RandomNum(column);

          if i >= domnum[lDomain] then
            rand := 0; // force item into FGroup 2 if half in FGroup 1
          if j >= domnum[lDomain] then
            rand := 1; // force item into FGroup 1 if half in FGroup 2
          if rand >= 0.50 then
            i := i + 1;
          if rand < 0.50 then
            j := j + 1;

          ticker := ticker + 1;
          frmMain.ProgressBar.Value := frmMain.ProgressBar.Value + 1;
          Application.ProcessMessages;

          for row := 1 to FNumberOfExaminees do
            if (UpCase(responsearray[row, column]) <> UpCase(FNAChar)) and
              (UpCase(responsearray[row, column]) <> UpCase(FOmitChar)) then
            begin // omits and nas are excluded
              FCh := responsearray[row, column];

              if rand >= 0.50 then
              begin
                if iteminfoarray[column, 3] = 'M' then
                  if UpCase(iteminfoarray[column, 1]) = UpCase(responsearray[row, column]) then
                    splitdomain1[row, 1 + cnum] := splitdomain1[row, 1 + cnum] + 1;

                if iteminfoarray[column, 3] <> 'M' then
                  splitdomain1[row, 1 + cnum] := splitdomain1[row, 1 + cnum] +
                    ConvertResponseToInteger(FCh, iteminfoarray, column, itemscalearray[column, 1]);
              end;

              if rand < 0.50 then
              begin
                if iteminfoarray[column, 3] = 'M' then
                  if UpCase(iteminfoarray[column, 1]) = UpCase(responsearray[row, column]) then
                    splitdomain2[row, 1 + cnum] := splitdomain2[row, 1 + cnum] + 1;

                if iteminfoarray[column, 3] <> 'M' then
                  splitdomain2[row, 1 + cnum] := splitdomain2[row, 1 + cnum] +
                    ConvertResponseToInteger(FCh, iteminfoarray, column, itemscalearray[column, 1]);
              end;

              if ticker <= domnum[lDomain] then
              begin
                if iteminfoarray[column, 3] = 'M' then
                  if UpCase(iteminfoarray[column, 1]) = UpCase(responsearray[row, column]) then
                    splitdomain1[row, 2 + cnum] := splitdomain1[row, 2 + cnum] + 1;

                if iteminfoarray[column, 3] <> 'M' then
                  splitdomain1[row, 2 + cnum] := splitdomain1[row, 2 + cnum] +
                    ConvertResponseToInteger(FCh, iteminfoarray, column, itemscalearray[column, 1]);
              end;

              if ticker > domnum[lDomain] then
              begin
                if iteminfoarray[column, 3] = 'M' then
                  if UpCase(iteminfoarray[column, 1]) = UpCase(responsearray[row, column]) then
                    splitdomain2[row, 2 + cnum] := splitdomain2[row, 2 + cnum] + 1;

                if iteminfoarray[column, 3] <> 'M' then
                  splitdomain2[row, 2 + cnum] := splitdomain2[row, 2 + cnum] +
                    ConvertResponseToInteger(FCh, iteminfoarray, column, itemscalearray[column, 1]);
              end;

              if (ticker mod 2) > 0 then
              begin
                if iteminfoarray[column, 3] = 'M' then
                  if UpCase(iteminfoarray[column, 1]) = UpCase(responsearray[row, column]) then
                    splitdomain1[row, 3 + cnum] := splitdomain1[row, 3 + cnum] + 1;

                if iteminfoarray[column, 3] <> 'M' then
                  splitdomain1[row, 3 + cnum] := splitdomain1[row, 3 + cnum] +
                    ConvertResponseToInteger(FCh, iteminfoarray, column, itemscalearray[column, 1]);
              end;

              if (ticker mod 2) = 0 then
              begin
                if iteminfoarray[column, 3] = 'M' then
                  if UpCase(iteminfoarray[column, 1]) = UpCase(responsearray[row, column]) then
                    splitdomain2[row, 3 + cnum] := splitdomain2[row, 3 + cnum] + 1;

                if iteminfoarray[column, 3] <> 'M' then
                  splitdomain2[row, 3 + cnum] := splitdomain2[row, 3 + cnum] +
                    ConvertResponseToInteger(FCh, iteminfoarray, column, itemscalearray[column, 1]);
              end;
            end; // for row loop
        end; // for column loop -- items loop
    end; // if FNumDomains > 1

  for row := 1 to FNumberOfExaminees do
  begin
    temp[1] := temp[1] + splitscore1[row, 1]; // split half - all items
    temp[2] := temp[2] + splitscore2[row, 1];
    temp[5] := temp[5] + splitscore1[row, 3]; // first-last - all items
    temp[6] := temp[6] + splitscore2[row, 3];
    temp[9] := temp[9] + splitscore1[row, 5]; // odd-even - all items
    temp[10] := temp[10] + splitscore2[row, 5];

    if FPreTest > 0 then
    begin
      temp[3] := temp[3] + splitscore1[row, 2];
      // split half - scored items only
      temp[4] := temp[4] + splitscore2[row, 2];
      temp[7] := temp[7] + splitscore1[row, 4];
      // first-last - scored items only
      temp[8] := temp[8] + splitscore2[row, 4];
      temp[11] := temp[11] + splitscore1[row, 6];
      // odd-even - scored items only
      temp[12] := temp[12] + splitscore2[row, 6];
    end;

    if FNumDomains > 1 then
      for lDomain := 1 to FNumDomains do
      begin
        cnum := 3 * (lDomain - 1);
        domsum[1, lDomain] := domsum[1, lDomain] + splitdomain1[row, 1 + cnum];
        domsum[2, lDomain] := domsum[2, lDomain] + splitdomain2[row, 1 + cnum];
        domsum[3, lDomain] := domsum[3, lDomain] + splitdomain1[row, 2 + cnum];
        domsum[4, lDomain] := domsum[4, lDomain] + splitdomain2[row, 2 + cnum];
        domsum[5, lDomain] := domsum[5, lDomain] + splitdomain1[row, 3 + cnum];
        domsum[6, lDomain] := domsum[6, lDomain] + splitdomain2[row, 3 + cnum];
      end;
  end;

  for v := 1 to 12 do
    tempM[v] := temp[v] / FNumberOfExaminees;

  if FNumDomains > 1 then
    for lDomain := 1 to FNumDomains do
      for v := 1 to 6 do
        domM[v, lDomain] := domsum[v, lDomain] / FNumberOfExaminees;

  for row := 1 to FNumberOfExaminees do
  begin
    sumsq[1] := sumsq[1] + (splitscore1[row, 1] - tempM[1]) * (splitscore1[row, 1] - tempM[1]);
    sumsq[2] := sumsq[2] + (splitscore2[row, 1] - tempM[2]) * (splitscore2[row, 1] - tempM[2]);
    sumsq[5] := sumsq[5] + (splitscore1[row, 3] - tempM[5]) * (splitscore1[row, 3] - tempM[5]);
    sumsq[6] := sumsq[6] + (splitscore2[row, 3] - tempM[6]) * (splitscore2[row, 3] - tempM[6]);
    sumsq[9] := sumsq[9] + (splitscore1[row, 5] - tempM[9]) * (splitscore1[row, 5] - tempM[9]);
    sumsq[10] := sumsq[10] + (splitscore2[row, 5] - tempM[10]) * (splitscore2[row, 5] - tempM[10]);

    if FPreTest > 0 then
    begin
      sumsq[3] := sumsq[3] + (splitscore1[row, 2] - tempM[3]) * (splitscore1[row, 2] - tempM[3]);
      sumsq[4] := sumsq[4] + (splitscore2[row, 2] - tempM[4]) * (splitscore2[row, 2] - tempM[4]);
      sumsq[7] := sumsq[7] + (splitscore1[row, 4] - tempM[7]) * (splitscore1[row, 4] - tempM[7]);
      sumsq[8] := sumsq[8] + (splitscore2[row, 4] - tempM[8]) * (splitscore2[row, 4] - tempM[8]);
      sumsq[11] := sumsq[11] + (splitscore1[row, 6] - tempM[11]) * (splitscore1[row, 6] - tempM[11]);
      sumsq[12] := sumsq[12] + (splitscore2[row, 6] - tempM[12]) * (splitscore2[row, 6] - tempM[12]);
    end;

    if FNumDomains > 1 then
      for lDomain := 1 to FNumDomains do
      begin
        cnum := 3 * (lDomain - 1);
        domsq[1, lDomain] := domsq[1, lDomain] + (splitdomain1[row, 1 + cnum] - domM[1, lDomain]) *
          (splitdomain1[row, 1 + cnum] - domM[1, lDomain]);
        domsq[2, lDomain] := domsq[2, lDomain] + (splitdomain2[row, 1 + cnum] - domM[2, lDomain]) *
          (splitdomain2[row, 1 + cnum] - domM[2, lDomain]);
        domsq[3, lDomain] := domsq[3, lDomain] + (splitdomain1[row, 2 + cnum] - domM[3, lDomain]) *
          (splitdomain1[row, 2 + cnum] - domM[3, lDomain]);
        domsq[4, lDomain] := domsq[4, lDomain] + (splitdomain2[row, 2 + cnum] - domM[4, lDomain]) *
          (splitdomain2[row, 2 + cnum] - domM[4, lDomain]);
        domsq[5, lDomain] := domsq[5, lDomain] + (splitdomain1[row, 3 + cnum] - domM[5, lDomain]) *
          (splitdomain1[row, 3 + cnum] - domM[5, lDomain]);
        domsq[6, lDomain] := domsq[6, lDomain] + (splitdomain2[row, 3 + cnum] - domM[6, lDomain]) *
          (splitdomain2[row, 3 + cnum] - domM[6, lDomain]);
      end;

    if FNumDomains > 1 then
      for lDomain := 1 to FNumDomains do
      begin
        cnum := 3 * (lDomain - 1);
        domcp[1, lDomain] := domcp[1, lDomain] + (splitdomain1[row, 1 + cnum] - domM[1, lDomain]) *
          (splitdomain2[row, 1 + cnum] - domM[2, lDomain]);
        domcp[2, lDomain] := domcp[2, lDomain] + (splitdomain1[row, 2 + cnum] - domM[3, lDomain]) *
          (splitdomain2[row, 2 + cnum] - domM[4, lDomain]);
        domcp[3, lDomain] := domcp[3, lDomain] + (splitdomain1[row, 3 + cnum] - domM[5, lDomain]) *
          (splitdomain2[row, 3 + cnum] - domM[6, lDomain]);
      end;

    CP1 := CP1 + (splitscore1[row, 1] - tempM[1]) * (splitscore2[row, 1] - tempM[2]);
    CP3 := CP3 + (splitscore1[row, 3] - tempM[5]) * (splitscore2[row, 3] - tempM[6]);
    CP5 := CP5 + (splitscore1[row, 5] - tempM[9]) * (splitscore2[row, 5] - tempM[10]);

    if FPreTest > 0 then
    begin
      CP2 := CP2 + (splitscore1[row, 2] - tempM[3]) * (splitscore2[row, 2] - tempM[4]);
      CP4 := CP4 + (splitscore1[row, 4] - tempM[7]) * (splitscore2[row, 4] - tempM[8]);
      CP6 := CP6 + (splitscore1[row, 6] - tempM[11]) * (splitscore2[row, 6] - tempM[12]);
    end;
  end; // for row to FNumberOfExaminees

  sd1 := sqrt(sumsq[1] / (FNumberOfExaminees - 1 + 1E-20));
  sd2 := sqrt(sumsq[2] / (FNumberOfExaminees - 1 + 1E-20));
  FTestStatsArray[2, 10] := CP1 / (sd1 * sd2 * (FNumberOfExaminees - 1) + 1E-20); // split half correlation
  FTestStatsArray[2, 13] := 2 * FTestStatsArray[2, 10] / (1 + FTestStatsArray[2, 10]); // spearman brown
  sd1 := sqrt(sumsq[5] / (FNumberOfExaminees - 1 + 1E-20));
  sd2 := sqrt(sumsq[6] / (FNumberOfExaminees - 1 + 1E-20));
  FTestStatsArray[2, 11] := CP3 / (sd1 * sd2 * (FNumberOfExaminees - 1) + 1E-20); // first-last
  FTestStatsArray[2, 14] := 2 * FTestStatsArray[2, 11] / (1 + FTestStatsArray[2, 11]); // spearman brown
  sd1 := sqrt(sumsq[9] / (FNumberOfExaminees - 1 + 1E-20));
  sd2 := sqrt(sumsq[10] / (FNumberOfExaminees - 1 + 1E-20));
  FTestStatsArray[2, 12] := CP5 / (sd1 * sd2 * (FNumberOfExaminees - 1) + 1E-20); // odd-even
  FTestStatsArray[2, 15] := 2 * FTestStatsArray[2, 12] / (1 + FTestStatsArray[2, 12]); // spearman brown

  if FPreTest > 0 then
  begin
    sd1 := sqrt(sumsq[3] / (FNumberOfExaminees - 1 + 1E-20));
    sd2 := sqrt(sumsq[4] / (FNumberOfExaminees - 1 + 1E-20));
    FTestStatsArray[1, 10] := CP2 / (sd1 * sd2 * (FNumberOfExaminees - 1) + 1E-20); // split half correlation
    FTestStatsArray[1, 13] := 2 * FTestStatsArray[1, 10] / (1 + FTestStatsArray[1, 10] + 1E-20); // spearman brown
    sd1 := sqrt(sumsq[7] / (FNumberOfExaminees - 1 + 1E-20));
    sd2 := sqrt(sumsq[8] / (FNumberOfExaminees - 1 + 1E-20));
    FTestStatsArray[1, 11] := CP4 / (sd1 * sd2 * (FNumberOfExaminees - 1) + 1E-20); // first-last
    FTestStatsArray[1, 14] := 2 * FTestStatsArray[1, 11] / (1 + FTestStatsArray[1, 11] + 1E-20); // spearman brown
    sd1 := sqrt(sumsq[11] / (FNumberOfExaminees - 1 + 1E-20));
    sd2 := sqrt(sumsq[12] / (FNumberOfExaminees - 1 + 1E-20));
    FTestStatsArray[1, 12] := CP6 / (sd1 * sd2 * (FNumberOfExaminees - 1) + 1E-20); // odd-even
    FTestStatsArray[1, 15] := 2 * FTestStatsArray[1, 12] / (1 + FTestStatsArray[1, 12] + 1E-20); // spearman brown
  end;

  if FNumDomains > 1 then
    for lDomain := 1 to FNumDomains do
    begin
      sd1 := sqrt(domsq[1, lDomain] / (FNumberOfExaminees - 1 + 1E-20));
      sd2 := sqrt(domsq[2, lDomain] / (FNumberOfExaminees - 1 + 1E-20));
      FDomainStatsArray[lDomain, 10] := domcp[1, lDomain] / (sd1 * sd2 * (FNumberOfExaminees - 1) + 1E-20);
      // split half correlation
      FDomainStatsArray[lDomain, 13] := 2 * FDomainStatsArray[lDomain, 10] /
        (1 + FDomainStatsArray[lDomain, 10] + 1E-20); // spearman brown
      sd1 := sqrt(domsq[3, lDomain] / (FNumberOfExaminees - 1 + 1E-20));
      sd2 := sqrt(domsq[4, lDomain] / (FNumberOfExaminees - 1 + 1E-20));
      FDomainStatsArray[lDomain, 11] := domcp[2, lDomain] / (sd1 * sd2 * (FNumberOfExaminees - 1) + 1E-20);
      // split half correlation
      FDomainStatsArray[lDomain, 14] := 2 * FDomainStatsArray[lDomain, 11] /
        (1 + FDomainStatsArray[lDomain, 11] + 1E-20); // spearman brown
      sd1 := sqrt(domsq[5, lDomain] / (FNumberOfExaminees - 1 + 1E-20));
      sd2 := sqrt(domsq[6, lDomain] / (FNumberOfExaminees - 1 + 1E-20));
      FDomainStatsArray[lDomain, 12] := domcp[3, lDomain] / (sd1 * sd2 * (FNumberOfExaminees - 1) + 1E-20);
      // split half correlation
      FDomainStatsArray[lDomain, 15] := 2 * FDomainStatsArray[lDomain, 12] /
        (1 + FDomainStatsArray[lDomain, 12] + 1E-20); // spearman brown
    end;
end;

procedure TItemanCalculation.CSem(itemscalearray: TIntegerArray; iteminfoarray, responsearray: TChar2Array;
  csemarray: TRealArray);
var
  row, column, multch: Integer;
  pss1, pvar1, Num, denom, k: Real;
begin
  pss1 := 0;
  multch := 0;
  FMinCsem := 0;
  FMaxCsem := -1;

  for column := 1 to FNumberOfItems do
    if (iteminfoarray[column, 2] = 'Y') and (iteminfoarray[column, 3] = 'M') then
    begin
      pss1 := pss1 + (FItemStatsArray[column, 2] - FTestStatsArray[1, 7]) *
        (FItemStatsArray[column, 2] - FTestStatsArray[1, 7]);
      multch := multch + 1;
    end;

  pvar1 := pss1 / (multch - 1);
  for row := 0 to round(FTestStatsArray[1, 1]) do
    csemarray[row, 1] := row * (FTestStatsArray[1, 1] - row) / (FTestStatsArray[1, 1] - 1); // CSem III for scored items

  Num := FScoredItems * (multch - 1) * pvar1;
  denom := FTestStatsArray[1, 2] * (multch - FTestStatsArray[1, 2]) - FTestStatsArray[1, 3] * FTestStatsArray[1, 3] -
    multch * (pvar1);
  k := Num / (denom + 1E-20);

  for row := 0 to round(FTestStatsArray[1, 1]) do
    csemarray[row, 2] := csemarray[row, 1] * (1 - k);

  for row := 0 to round(FTestStatsArray[1, 1]) do
    if sqrt(abs(csemarray[row, 2])) > FMaxCsem then
      FMaxCsem := sqrt(abs(csemarray[row, 2]));
end;

function TItemanCalculation.RandomNum(counter: Integer): Real;
var
  ly, lz, ltrxyz: LongInt;
  lx, lxyz: Real;
begin
  FIX := trunc(171 * (FIX mod 177) - 2 * (FIX / 177));
  FIY := trunc(172 * (FIY mod 176) - 35 * (FIY / 176));
  FIZ := trunc(170 * (FIZ mod 178) - 63 * (FIZ / 178));

  if FIX < 0.001 then
    FIX := FIX + 30269;
  if FIY < 0.001 then
    FIY := FIY + 30307;
  if FIZ < 0.001 then
    FIZ := FIZ + 30323;

  lx := FIX;
  ly := FIY;
  lz := FIZ;
  lxyz := lx / 30329.0 + ly / 30307.0 + lz / 30323.0;
  ltrxyz := trunc(lxyz);
  Result := abs(lxyz - ltrxyz);
end;

procedure TItemanCalculation.ShellSortInt(numericarray: TIntegerArray; n: Integer);
label
  Jumpout;

var
  h, i, j, Value: Integer;
begin
  { find the most efficient value for h }
  h := 1;

  repeat
    h := (3 * h) + 1;
  until (h > n);

  { shell sort }
  repeat
    h := h div 3;

    for i := succ(h) to n do
    begin
      Value := numericarray[i, 0];
      j := i;

      while (numericarray[j - h, 0] > Value) do
      begin
        numericarray[j, 0] := numericarray[j - h, 0];
        j := j - h;

        if j <= h then
          goto Jumpout;
      end;

    Jumpout:
      numericarray[j, 0] := Value;
    end;
  until (h = 1);
end;

procedure TItemanCalculation.RawFreq(rawfrequency, scorearray: TIntegerArray);
var
  lRow, lColumn, lIJ: Integer;
begin
  for lRow := 0 to round(FTestStatsArray[1, 5] - FTestStatsArray[1, 4]) do
    for lColumn := 1 to 3 do
      rawfrequency[lRow, lColumn] := 0;

  lIJ := 0;
  FNumUniqueFreq := 1;

  for lRow := 1 to FNumberOfExaminees do
  begin
    if lRow = 1 then
      rawfrequency[lIJ, 1] := scorearray[lRow, 0]; // min score

    if rawfrequency[lIJ, 1] = scorearray[lRow, 0] then
    begin
      inc(rawfrequency[lIJ, 2]); // increment observed frequency
      inc(rawfrequency[lIJ, 3]); // increment cumulative frequency
    end;

    if rawfrequency[lIJ, 1] < scorearray[lRow, 0] then
    begin
      inc(lIJ);
      inc(FNumUniqueFreq);

      rawfrequency[lIJ, 1] := scorearray[lRow, 0];
      inc(rawfrequency[lIJ, 2]); // increment observed frequency
      rawfrequency[lIJ, 3] := rawfrequency[lIJ - 1, 3]; //
      inc(rawfrequency[lIJ, 3]); // increment cumulative frequency
    end;
  end;
end;

procedure TItemanCalculation.BBOindex(iteminfoarray, responsearray: TChar2Array; itemscalearray: TIntegerArray);
var
  n: Integer;
  // Same as H-H Errors In Common - need to sum before final calculation
  k: Integer;
  // Same as H-H Exact Errors In Common - need to sum before final calculation
  k_step: Integer; // For cycling sum from k to N
  lOption, exam1, exam2, item, w1, w2, w3: Integer;
  r1, r2: Char;
  bb, temp_p, sum_p, BBO_flag_threshold: Real;
  PforBBArray: Array of Real; // For calculating average P
  lRow, lColumn: Integer;
  lBBO_P: Real;
begin
  SetLength(FEICArray, succ(FNumberOfExaminees), (FNumberOfExaminees + 1));
  SetLength(FEEICArray, succ(FNumberOfExaminees), (FNumberOfExaminees + 1));
  SetLength(PforBBArray, succ(FNumberOfItems));
  BBO_flag_threshold := StrToFloatDef(frmMain.BBO_flag_box.Text, 0);

  { ---Determine P based on observed distractor proportions--- }
  sum_p := 0;
  for lRow := 1 to FNumberOfItems do
    if (iteminfoarray[lRow, 2] = 'Y') then
    begin
      PforBBArray[lRow] := 0;

      for lOption := 1 to FMaxItemOptions do
        if lOption <> KeyAsInt(iteminfoarray, lRow) then
        begin
          temp_p := FOptionPropsArray[lRow, lOption] / (1 - FItemStatsArray[lRow, 2] + 0.00000000001);
          // Divide by (1- ItemP) to get prop of incorrect responses only
          PforBBArray[lRow] := PforBBArray[lRow] + temp_p * temp_p;
          sum_p := sum_p + temp_p * temp_p;
        end;
    end; // initialize P array

  lBBO_P := sum_p / FScoredItems;

  { ----- Count EIC and EEIC for all pairs ----- }
  for exam1 := 1 to FNumberOfExaminees do { For examinee 1 to N, }
    for exam2 := 1 to FNumberOfExaminees do { Compare examinee 1 to N }
    begin
      FEICArray[exam1, exam2] := 0;
      FEEICArray[exam1, exam2] := 0;
      FBBOindexArray[exam1, exam2] := 0;
    end;

  for exam1 := 1 to FNumberOfExaminees do { For examinee 1 to N, }
    for exam2 := 1 to FNumberOfExaminees do { Compare examinee 1 to N }
      if exam1 > exam2 then // for lower triangular values
      begin
        for item := 1 to FNumberOfItems do
          if (iteminfoarray[item, 2] = 'Y') and
          // exclude 'N' and 'P' items
            (responsearray[exam1, item] <> FOmitChar) and (responsearray[exam1, item] <> FNAChar) and
            (responsearray[exam2, item] <> FOmitChar) and (responsearray[exam2, item] <> FNAChar)
          // exclude omit and na items
          then
          begin
            r1 := responsearray[exam1, item];
            { response from examinee 1 for item i }
            r2 := responsearray[exam2, item];
            { response from examinee 2 for item i }

            { BBOindex is calculated for multiple choice items }
            if (r1 <> iteminfoarray[item, 1]) then // and (ItemType[item] = 'M') then
              w1 := 1
            else
              w1 := 0; { If the response is incorrect, weight 1 for examinee 1 }

            if (r2 <> iteminfoarray[item, 1]) then // and (ItemType[item] = 'M') then
              w2 := 1
            else
              w2 := 0; { If the response is incorrect, weight 1 for examinee 2 }

            If (r1 = r2) then // and (ItemKey[item] = 'M') then
              w3 := 1
            else
              w3 := 0; { If the response are the same, weight 1 }

            FEICArray[exam1, exam2] := (FEICArray[exam1, exam2] + (w1 * w2));
            // count 1 if both incorrect
            FEEICArray[exam1, exam2] := (FEEICArray[exam1, exam2] + (w1 * w3));
            // count 1 if both exactly incorrect
          end;

        // BB index is loop from 18 to 20 as in example...
        // Do we want P of 18 common, or of 18 or more common?
        // Cizek and Khalid have N and k switched!
        k := FEEICArray[exam1, exam2];
        n := FEICArray[exam1, exam2];

        for k_step := k to n do
        begin
          bb := Factorial(n) / (Factorial(k_step) * Factorial(n - k_step) + 0.0000000000000001) * power(lBBO_P, k_step)
            * power((1 - lBBO_P), (n - k_step));
          FBBOindexArray[exam1, exam2] := FBBOindexArray[exam1, exam2] + bb;
        end; // k_step
      end; // end if exam1>exam2

  { Flagging for B&B index procedure }
  for lRow := 1 to FNumberOfExaminees do
    FBBOFlags[lRow] := 0;

  for lRow := 1 to FNumberOfExaminees do
  begin
    for lColumn := 1 to FNumberOfExaminees do
    begin
      if lRow > lColumn then
      begin
        bb := FBBOindexArray[lRow, lColumn];

        if bb < BBO_flag_threshold then
        begin
          FBBOFlags[lRow] := FBBOFlags[lRow] + 1;
          FBBOFlags[lColumn] := FBBOFlags[lColumn] + 1;
        end;
      end;
    end; // end lColumn
  end;
end;

function TItemanCalculation.Factorial(Num: LongInt): LongInt;
var
  i, j: LongInt;
begin
  Result := 0;
  j := round(abs(Num));

  if j = 1 then
    Result := j;
  if j = 0 then
    Result := 1;

  if j >= 2 then
  begin
    Result := j;

    for i := j - 1 downto 2 do
      Result := i * Result;
  end;
end;

function TItemanCalculation.KeyAsInt(iteminfoarray: TChar2Array; item: Integer): Integer;
begin
  Result := 0;
  if iteminfoarray[item, 1] = 'A' then
    Result := 1;
  if iteminfoarray[item, 1] = 'B' then
    Result := 2;
  if iteminfoarray[item, 1] = 'C' then
    Result := 3;
  if iteminfoarray[item, 1] = 'D' then
    Result := 4;
  if iteminfoarray[item, 1] = 'E' then
    Result := 5;
  if iteminfoarray[item, 1] = 'F' then
    Result := 6;
  if iteminfoarray[item, 1] = 'G' then
    Result := 7;
  if iteminfoarray[item, 1] = 'H' then
    Result := 8;
  if iteminfoarray[item, 1] = 'I' then
    Result := 9;
  if iteminfoarray[item, 1] = '0' then
    Result := 0;
  if iteminfoarray[item, 1] = '1' then
    Result := 1;
  if iteminfoarray[item, 1] = '2' then
    Result := 2;
  if iteminfoarray[item, 1] = '3' then
    Result := 3;
  if iteminfoarray[item, 1] = '4' then
    Result := 4;
  if iteminfoarray[item, 1] = '5' then
    Result := 5;
  if iteminfoarray[item, 1] = '6' then
    Result := 6;
  if iteminfoarray[item, 1] = '7' then
    Result := 7;
  if iteminfoarray[item, 1] = '8' then
    Result := 8;
  if iteminfoarray[item, 1] = '9' then
    Result := 9;
end;

function TItemanCalculation.Livingston: Real;
var
  zscore, pval, xbar, sd1, var1, Num, denom, square: Real;
begin
  with frmMain do
  begin
    rawcut := classcutbox.Text.ToDouble;
    if ScoredClass.IsChecked then
      pval := classcutbox.Text.ToDouble / FTestStatsArray[1, 1];
    if ScaledClass.IsChecked then
    begin
      if standscale.IsChecked then
      begin
        zscore := (classcutbox.Text.ToDouble - newmeanbox.Text.ToDouble) / newsdbox.Text.ToDouble;
        rawcut := (zscore * FTestStatsArray[1, 3]) + FTestStatsArray[1, 2];
        pval := rawcut / FTestStatsArray[1, 1];
      end;
      if linearscale.IsChecked then
      begin
        xbar := (FTestStatsArray[1, 2] + interceptbox.Text.ToDouble) * slopebox.Text.ToDouble;
        sd1 := FTestStatsArray[1, 3] * slopebox.Text.ToDouble;
        zscore := (classcutbox.Text.ToDouble - xbar) / sd1;
        rawcut := zscore * FTestStatsArray[1, 3] + FTestStatsArray[1, 2];
        pval := rawcut / FTestStatsArray[1, 1];
      end;
    end; // if scaledclass checked!
    if pval <= 0 then
      pval := 0.00000001;
    if pval >= 1.0 then
      pval := 0.999999999;
    var1 := FTestStatsArray[1, 3] * FTestStatsArray[1, 3];
    square := power(FTestStatsArray[1, 2] - (FTestStatsArray[1, 1] * pval), 2);
    Num := var1 * FTestStatsArray[1, 6] + square;
    denom := var1 + square;
    Result := Num / denom;
  end;
end;

procedure TItemanCalculation.SubGroups(aItemInfoArray, aResponseArray: TChar2Array; aOmch: Char; aI: Integer;
  aRankIngArray, aReduceArray, aItemsCaleArray: TIntegerArray);
var
  lA, lRow, lColumn: Integer;
  jfor: Integer;
begin
  for lA := 1 to 105 do
    FSubGroupPArray[lA] := 0;

  // --- Calculate N for SubGroups ---
  if aItemInfoArray[aI, 2] <> 'N' then
    for lRow := 1 to FNumberOfExaminees do
      if aResponseArray[lRow, aI] <> FOmitChar then
        if aResponseArray[lRow, aI] <> FNAChar then
        begin
          if aItemInfoArray[aI, 3] = 'P' then
            FWts := 1
          else
            FWts := 0;

          FGroup := aRankIngArray[lRow, 2];
          FCh := aResponseArray[lRow, aI];
          FResponseAsInt := ConvertResponseToInteger(FCh, aItemInfoArray, aI, aItemsCaleArray[aI, 1]);
          FSubGroupPArray[((FResponseAsInt - 1 + FWts) * FNumCut + FGroup)] :=
            FSubGroupPArray[((FResponseAsInt - 1 + FWts) * FNumCut + FGroup)] + 1;
          { each lRow,lColumn is an item: LMH for lA, LMH for B... 1 is the high FGroup }
        end;

  { --- Convert all those subgroup N to P --- }
  if aItemInfoArray[aI, 2] <> 'N' then
  begin
    lColumn := 1;

    if aItemInfoArray[aI, 3] = 'P' then
      FWts := 1
    else
      FWts := 0;

    while lColumn < (FMaxItemOptions * FNumCut) do
    begin
      for jfor := 1 to FNumCut do
        FSubGroupPArray[lColumn + pred(jfor)] := FSubGroupPArray[lColumn + pred(jfor)] /
          (FGroupCounts[jfor] - aReduceArray[aI, jfor] + 1E-20);

      lColumn := lColumn + FNumCut;
    end;
  end;
end;

function TItemanCalculation.GetFloatValue(aFloatValueInStr: string = '0'; aFloatDidgets: Integer = -3): Extended;
var
  lFloatChar: String;
begin
  if (FormatSettings.DecimalSeparator = '.') then
    lFloatChar := ','
  else
    lFloatChar := '.';

  Result := RoundTo(StrToFloatDef(StringReplace(aFloatValueInStr, lFloatChar, FormatSettings.DecimalSeparator,
    [rfReplaceAll]), 0), aFloatDidgets);
end;

function TItemanCalculation.GetColumnsCountOfCSVFile(Filename: string): Integer;
var
  F: TextFile;
  SL: TStringList;
  Line: String;
begin
  AssignFile(F, Filename);
  Reset(F);
  ReadLn(F, Line);
  CloseFile(F);

  SL := TStringList.Create;
  SL.StrictDelimiter := True;
  try
    SL.DelimitedText := Line;
    Result := SL.count;
  finally
    SL.Free;
  end;
end;

{ TBitmapHelper }

procedure TBitmapHelper.LoadFromResource(AResourceName: String; AWidth: Integer; AHeight: Integer);
var
  ResourceStream: TResourceStream;
  LKoeff: double;
begin
  ResourceStream := TResourceStream.Create(HInstance, AResourceName, RT_RCDATA);
  try
    LKoeff := 1;
    SetSize(0, 0);
    Self.LoadFromStream(ResourceStream);
    if (AWidth > 0) and (AHeight > 0) then
    begin
      Resize(AWidth, AHeight);
      Exit;
    end
    else if (AHeight = 0) and (AWidth > 0) then
      LKoeff := AWidth / Width
    else if (AHeight > 0) and (AWidth = 0) then
      LKoeff := AHeight / Height;
    if LKoeff < 1 then
      Resize(ceil(Width * LKoeff), ceil(Height * LKoeff));
  finally
    FreeAndNil(ResourceStream);
  end;
end;

end.
