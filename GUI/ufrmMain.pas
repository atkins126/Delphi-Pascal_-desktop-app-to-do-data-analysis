unit ufrmMain;

interface

uses
  System.SysUtils,
  System.Types,
  System.UITypes,
  System.Classes,
  System.Variants,

  System.Rtti,
  vkbdhelper,
  FMX.DialogService,
  System.IOUtils,
  System.DateUtils,
  System.Math,
  ufrmMultipleRunsFile,
  uProtect,
  uItemanCalculation,
  uItemanExport,
  StrUtils,

{$REGION ' POPUP WIN '}
  ufrmLicense,
{$ENDREGION ' POPUP WIN '}
{$IF DEFINED(Win64) or DEFINED(Win32)}
  WinAPI.ShellApi,
  WinAPI.Windows,
{$ENDIF}
  FMX.Types,
  FMX.Controls,
  FMX.Forms,
  FMX.Graphics,
  FMX.Dialogs,
  FMX.Layouts,
  FMX.Controls.Presentation,
  FMX.StdCtrls,
  FMX.ListBox,
  FMX.Objects,
  FMX.Effects,
  FMX.Colors,
  FMX.MultiView,
  FMX.TabControl,
  FMX.Edit,
  FMX.Ani,
  FMX.ScrollBox,
  FMX.Memo,
  FMX.Memo.Types,
  FMX.EditBox,
  FMX.SpinBox,
  FMX.NumberBox;

type
  TPWideCharArray = array of PWideChar;

  TWndProc = function(hwnd: hwnd; uMsg: UINT; wParam: wParam; lParam: lParam): LRESULT; stdcall;

  TMyTask = class(TThread)
  private const
    ItemControlVariableAdditionalRows = 1;
    function GetDataMatrixVariableAdditionalColumns: Integer;
  protected
    FOnTaskStarted: TThreadMethod;
    FOnTaskFinished: TThreadMethod;
    FOnTaskEnded: TThreadMethod;
    FOnTaskException: TThreadMethod;

    Fdata: string;
    Fctrl: string;
    Foutp: string;
    Fmrf1: string;

    procedure DoStarted();
    procedure DoFinished();
    procedure DoEnded();
    procedure DoException();
  public
    procedure Execute(); override;

    property OnTaskStarted: TThreadMethod read FOnTaskStarted write FOnTaskStarted;
    property OnTaskFinished: TThreadMethod read FOnTaskFinished write FOnTaskFinished;
    property OnTaskEnded: TThreadMethod read FOnTaskEnded write FOnTaskEnded;
    property OnTaskException: TThreadMethod read FOnTaskException write FOnTaskException;

    property InputData: string read Fdata write Fdata;
    property InputCtrl: string read Fctrl write Fctrl;
    property InputOutp: string read Foutp write Foutp;
    property InputMrf1: string read Fmrf1 write Fmrf1;
  end;

  TfrmMain = class(TForm)
    StyleBook1: TStyleBook;
    LayoutContent: TLayout;
    tbMain: TToolBar;
    btnMenu: TButton;
    mtvMenu: TMultiView;
    LayoutMenuContent: TLayout;
    gplMain: TGridPanelLayout;
    Layout12: TLayout;
    Layout13: TLayout;
    tcMain: TTabControl;
    TabItem1: TTabItem;
    Layout14: TLayout;
    TabItem2: TTabItem;
    Layout15: TLayout;
    TabItem3: TTabItem;
    Layout16: TLayout;
    TabItem4: TTabItem;
    Layout17: TLayout;
    Layout18: TLayout;
    gplTopInfoBlock: TGridPanelLayout;
    gplTab1: TGridPanelLayout;
    Layout19: TLayout;
    Panel2: TPanel;
    Text1: TText;
    Layout20: TLayout;
    Layout21: TLayout;
    Layout22: TLayout;
    Layout23: TLayout;
    Layout24: TLayout;
    Layout25: TLayout;
    gplTab1Row1: TGridPanelLayout;
    Layout26: TLayout;
    Text2: TText;
    Layout27: TLayout;
    eDataMatrixFile: TEdit;
    GridPanelLayout1: TGridPanelLayout;
    Layout28: TLayout;
    Text3: TText;
    Layout29: TLayout;
    eItemControlFile: TEdit;
    GridPanelLayout2: TGridPanelLayout;
    Layout30: TLayout;
    Text4: TText;
    Layout31: TLayout;
    eOutputFile: TEdit;
    GridPanelLayout3: TGridPanelLayout;
    Layout32: TLayout;
    Text5: TText;
    Layout33: TLayout;
    RunTitleBox: TEdit;
    GridPanelLayout4: TGridPanelLayout;
    Layout34: TLayout;
    Layout35: TLayout;
    eExternalScoreFile: TEdit;
    cbExternalScoreFile: TCheckBox;
    cbDataMatrixFile: TCheckBox;
    Layout36: TLayout;
    gplTab2: TGridPanelLayout;
    Layout52: TLayout;
    Panel4: TPanel;
    Text12: TText;
    Layout53: TLayout;
    GridPanelLayout6: TGridPanelLayout;
    Layout54: TLayout;
    Layout55: TLayout;
    Layout56: TLayout;
    Layout57: TLayout;
    Layout60: TLayout;
    Layout63: TLayout;
    Layout66: TLayout;
    GridPanelLayout5: TGridPanelLayout;
    Layout58: TLayout;
    Layout59: TLayout;
    GridPanelLayout7: TGridPanelLayout;
    Layout61: TLayout;
    Layout62: TLayout;
    GridPanelLayout8: TGridPanelLayout;
    Layout69: TLayout;
    Layout70: TLayout;
    Panel5: TPanel;
    Text13: TText;
    Panel6: TPanel;
    Text14: TText;
    Text15: TText;
    Text17: TText;
    Text18: TText;
    Layout71: TLayout;
    Layout72: TLayout;
    Layout73: TLayout;
    Button1: TButton;
    Button2: TButton;
    eTab2Numbers1: TEdit;
    Button3: TButton;
    IDColumnsBegin: TEdit;
    Button4: TButton;
    Button5: TButton;
    itemColumnBox: TEdit;
    Button6: TButton;
    DelimitBox: TCheckBox;
    IncludeIDBox: TCheckBox;
    gplTab2Radios: TGridPanelLayout;
    Layout74: TLayout;
    Layout75: TLayout;
    CommaBox: TRadioButton;
    TabBox: TRadioButton;
    GridPanelLayout9: TGridPanelLayout;
    Layout64: TLayout;
    Layout65: TLayout;
    Text16: TText;
    GridPanelLayout10: TGridPanelLayout;
    Layout67: TLayout;
    Text19: TText;
    Layout68: TLayout;
    OmitCharBox: TEdit;
    NACharBox: TEdit;
    Layout76: TLayout;
    Layout77: TLayout;
    Panel7: TPanel;
    Text20: TText;
    Layout78: TLayout;
    DifCheckBox: TCheckBox;
    Layout82: TLayout;
    GridPanelLayout12: TGridPanelLayout;
    Layout83: TLayout;
    Text22: TText;
    Layout84: TLayout;
    Button9: TButton;
    Button10: TButton;
    NumDifGroups: TEdit;
    Layout79: TLayout;
    GridPanelLayout11: TGridPanelLayout;
    Layout80: TLayout;
    Text21: TText;
    Layout81: TLayout;
    Button7: TButton;
    Button8: TButton;
    GroupColumn: TEdit;
    Layout85: TLayout;
    GridPanelLayout13: TGridPanelLayout;
    Layout86: TLayout;
    Text23: TText;
    Layout88: TLayout;
    Layout91: TLayout;
    Group1Code: TEdit;
    Layout87: TLayout;
    Text25: TText;
    Layout92: TLayout;
    Group2Code: TEdit;
    GridPanelLayout14: TGridPanelLayout;
    Layout89: TLayout;
    Text24: TText;
    Layout90: TLayout;
    Group1label: TEdit;
    Layout93: TLayout;
    Text26: TText;
    Layout94: TLayout;
    Group2label: TEdit;
    gplTab3: TGridPanelLayout;
    Layout95: TLayout;
    Panel8: TPanel;
    Text27: TText;
    Layout96: TLayout;
    Layout99: TLayout;
    Layout103: TLayout;
    Layout109: TLayout;
    Layout113: TLayout;
    Layout116: TLayout;
    Layout119: TLayout;
    Layout120: TLayout;
    Panel11: TPanel;
    Text35: TText;
    Layout121: TLayout;
    ClassCheckBox: TCheckBox;
    Layout122: TLayout;
    Layout125: TLayout;
    Layout128: TLayout;
    GridPanelLayout25: TGridPanelLayout;
    Layout129: TLayout;
    Text38: TText;
    Layout130: TLayout;
    ClassCutBox: TEdit;
    Layout133: TLayout;
    GridPanelLayout26: TGridPanelLayout;
    Layout134: TLayout;
    Text40: TText;
    Layout135: TLayout;
    LowLabel: TEdit;
    Layout136: TLayout;
    Text41: TText;
    Layout137: TLayout;
    HighLabel: TEdit;
    Text28: TText;
    Layout102: TLayout;
    ScaledCheckBox: TCheckBox;
    Layout97: TLayout;
    ScaledDomain: TCheckBox;
    Text29: TText;
    GridPanelLayout15: TGridPanelLayout;
    Layout98: TLayout;
    Layout100: TLayout;
    Layout101: TLayout;
    Layout104: TLayout;
    LineARScale: TRadioButton;
    Layout105: TLayout;
    InterceptBox: TEdit;
    Text31: TText;
    SlopeBox: TEdit;
    Text30: TText;
    GridPanelLayout16: TGridPanelLayout;
    Layout106: TLayout;
    StandScale: TRadioButton;
    Layout107: TLayout;
    Text32: TText;
    Layout108: TLayout;
    NewMeanBox: TEdit;
    Layout110: TLayout;
    Text33: TText;
    Layout111: TLayout;
    NewSDBox: TEdit;
    ScoredClass: TRadioButton;
    ScaledClass: TRadioButton;
    gplMainForTab4: TGridPanelLayout;
    Layout112: TLayout;
    Layout114: TLayout;
    gplFlags: TGridPanelLayout;
    Layout115: TLayout;
    Panel9: TPanel;
    Text34: TText;
    Layout117: TLayout;
    Layout118: TLayout;
    Layout124: TLayout;
    Layout127: TLayout;
    Layout131: TLayout;
    Layout142: TLayout;
    Layout148: TLayout;
    Layout149: TLayout;
    GridPanelLayout20: TGridPanelLayout;
    Layout154: TLayout;
    Text46: TText;
    Layout155: TLayout;
    KeyFlag: TEdit;
    GridPanelLayout17: TGridPanelLayout;
    Layout123: TLayout;
    Text36: TText;
    Layout126: TLayout;
    DifFlag: TEdit;
    GridPanelLayout22: TGridPanelLayout;
    Layout161: TLayout;
    Layout162: TLayout;
    Text45: TText;
    Layout163: TLayout;
    Text49: TText;
    GridPanelLayout23: TGridPanelLayout;
    Layout164: TLayout;
    Layout165: TLayout;
    Layout166: TLayout;
    GridPanelLayout18: TGridPanelLayout;
    Layout132: TLayout;
    Layout138: TLayout;
    Layout139: TLayout;
    GridPanelLayout19: TGridPanelLayout;
    Layout140: TLayout;
    Layout141: TLayout;
    Layout143: TLayout;
    Text37: TText;
    Text39: TText;
    Text42: TText;
    LowPFlag: TEdit;
    LowRFlag: TEdit;
    LowMeanFlag: TEdit;
    HighPFlag: TEdit;
    HighRFlag: TEdit;
    HighMeanFlag: TEdit;
    GridPanelLayout21: TGridPanelLayout;
    Layout144: TLayout;
    Text43: TText;
    Layout145: TLayout;
    BBO_flag_box: TEdit;
    gplTab4Main: TGridPanelLayout;
    Layout146: TLayout;
    Panel10: TPanel;
    Text44: TText;
    Layout150: TLayout;
    Layout152: TLayout;
    Layout156: TLayout;
    Layout157: TLayout;
    Layout169: TLayout;
    Layout175: TLayout;
    Layout176: TLayout;
    Panel12: TPanel;
    Text54: TText;
    Layout177: TLayout;
    Layout178: TLayout;
    Layout179: TLayout;
    Layout180: TLayout;
    GridPanelLayout24: TGridPanelLayout;
    Layout151: TLayout;
    Text47: TText;
    Layout189: TLayout;
    Text58: TText;
    Layout147: TLayout;
    Button11: TButton;
    Button12: TButton;
    Layout188: TLayout;
    Button13: TButton;
    Button14: TButton;
    GridPanelLayout31: TGridPanelLayout;
    Layout153: TLayout;
    Text48: TText;
    Layout190: TLayout;
    Text59: TText;
    Layout191: TLayout;
    Button15: TButton;
    Button16: TButton;
    Layout192: TLayout;
    Button17: TButton;
    Button18: TButton;
    GridPanelLayout32: TGridPanelLayout;
    Layout193: TLayout;
    Text60: TText;
    Layout194: TLayout;
    Text61: TText;
    Layout195: TLayout;
    Button19: TButton;
    Button20: TButton;
    Layout196: TLayout;
    Button21: TButton;
    Button22: TButton;
    OmitAsinc: TCheckBox;
    SpurBox: TCheckBox;
    GridPanelLayout27: TGridPanelLayout;
    Layout158: TLayout;
    Text50: TText;
    Layout159: TLayout;
    Text51: TText;
    Layout160: TLayout;
    Button23: TButton;
    Button24: TButton;
    Precision: TEdit;
    PlotsBox: TCheckBox;
    TableBox: TCheckBox;
    GridPanelLayout28: TGridPanelLayout;
    Layout167: TLayout;
    Text52: TText;
    Layout168: TLayout;
    Text53: TText;
    Layout170: TLayout;
    Button25: TButton;
    Button26: TButton;
    CutPoint: TEdit;
    Layout171: TLayout;
    CollusionBox: TCheckBox;
    Layout172: TLayout;
    SaveControl: TCheckBox;
    Layout173: TLayout;
    ScoredMatrixBox: TCheckBox;
    Layout174: TLayout;
    IncLomit: TCheckBox;
    Layout181: TLayout;
    IncLnr: TCheckBox;
    LayoutSplash: TLayout;
    PanelSplashBG: TPanel;
    LayoutSplashContainer: TLayout;
    gplSplash: TGridPanelLayout;
    pSplashContent: TPanel;
    FloatAnimation1: TFloatAnimation;
    SaveDialog: TSaveDialog;
    OpenDialog: TOpenDialog;
    GoodLogo: TImage;
    LayoutMenu: TLayout;
    pMenuBG: TPanel;
    gplMenu: TGridPanelLayout;
    Layout1: TLayout;
    pFirstGroup: TPanel;
    tFileMenuGroup: TText;
    Layout2: TLayout;
    Layout3: TLayout;
    Layout4: TLayout;
    Layout5: TLayout;
    Layout6: TLayout;
    Layout7: TLayout;
    Layout8: TLayout;
    Layout9: TLayout;
    Layout10: TLayout;
    Layout11: TLayout;
    Layout44: TLayout;
    Panel3: TPanel;
    Text9: TText;
    Layout45: TLayout;
    gplLicenceRow1: TGridPanelLayout;
    Layout47: TLayout;
    Text10: TText;
    Layout48: TLayout;
    tLicenceDaysLeft: TText;
    Layout46: TLayout;
    gplLicenceRow2: TGridPanelLayout;
    Layout49: TLayout;
    Text11: TText;
    Layout50: TLayout;
    tLicenceExpiresDate: TText;
    Layout51: TLayout;
    btnLicense: TButton;
    btnTabFiles: TButton;
    btnInputFormat: TButton;
    btnScoringOptions: TButton;
    btnOutputOptions: TButton;
    Panel1: TPanel;
    tTabsMenu: TText;
    btnSaveTheProgramDefaults: TButton;
    btnCreateMultipleRunsFile: TButton;
    btnRunSavedMultipleRunsFile: TButton;
    btnOpenAnOptionsFile: TButton;
    btnSaveAnOptionsFile: TButton;
    refill: TImageControl;
    popPrecisionList: TPopup;
    pPrecisionListBG: TPanel;
    lbPrecision: TListBox;
    popCutPointList: TPopup;
    pCutPointListBG: TPanel;
    lbCutPoint: TListBox;
    Button27: TButton;
    Button28: TButton;
    Button29: TButton;
    Button30: TButton;
    pnlLoad: TPanel;
    LayoutLoading: TLayout;
    LayoutAnime: TLayout;
    AniIndicator1: TAniIndicator;
    Layout182: TLayout;
    ProgressLabel: TText;
    mLog: TMemo;
    ProgressBar: TProgressBar;
    btnRun: TButton;
    MaxRpbisBox: TNumberBox;
    MinPBox: TNumberBox;
    MaxPBox: TNumberBox;
    MinMeanBox: TNumberBox;
    MaxMeanBox: TNumberBox;
    MinRpbisBox: TNumberBox;
    procedure btnMenuClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnMenuItemOnClick(Sender: TObject);
    procedure btnTab1OnClick(Sender: TObject);
    procedure btnTab2NumbersOnClick(Sender: TObject);
    procedure DelimitBoxChange(Sender: TObject);
    procedure cbExternalScoreFileChange(Sender: TObject);
    procedure DifCheckBoxChange(Sender: TObject);
    procedure ScaledCheckBoxChange(Sender: TObject);
    procedure ClassCheckBoxChange(Sender: TObject);

{$REGION ' POPUP WIN '}
    procedure OpenLicensePad(Sender: TObject);

    procedure btnSplashClick(Sender: TObject);
    procedure PanelSplashBGClick(Sender: TObject);
{$ENDREGION ' POPUP WIN '}
    procedure ScoredMatrixBoxChange(Sender: TObject);
    procedure btnRunClick(Sender: TObject);
    procedure ScoredClassClick(Sender: TObject);
    procedure lbPrecisionItemClick(const Sender: TCustomListBox; const Item: TListBoxItem);
    procedure lbCutPointItemClick(const Sender: TCustomListBox; const Item: TListBoxItem);
    procedure FormResize(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure eTab2Numbers1Change(Sender: TObject);
    procedure StandScaleChange(Sender: TObject);
  private
    { Thread }
    FMyTask: TMyTask;
    FTreadExceptMessage: string;
    // l:TCanvas;

    FMRFFile: string;
    FAPPPath: string;
    FIs_Demo: Boolean;
    FMRFText1: string;
    FDocxOutput: string;
    FUseMRF: Boolean;
    // FOutputDir: string;

    FProtect: TProtect;
    FFirstActivate: Boolean;
    FInst: Cardinal;

    ReadFlags: Array [1 .. 50] of Char;
    FlagArray: Array [1 .. 8] of string;

    FFormatSettings: TFormatSettings;

    procedure CheckOptionValue(const AEdit: TEdit);
    function SetValueInRange<T>(AValue: T; AMin: T; AMax: T): T;

{$REGION ' POPUP WIN '}
    procedure SplashShower(aMode: Boolean = False; aDefaultHeight: Integer = 300; aSplashWidth: Integer = 300;
      aPanelBGStyleName: string = 'PanelSplashGrayStyle');
{$ENDREGION ' POPUP WIN '}
    function GetFloatValue(aFloatValueInStr: string = '0'; aFloatDidgets: Integer = -3): Extended;

    procedure RequestLicenseKey();
    procedure ConfirmLicense();

    procedure UpdateLicenseDesc();
    procedure SaveOptionToFile(aFileNameWithPath: string = '');

    procedure GetBuildInfo(const SourceFile: String; var V1, V2, V3, V4: Word);
    function GetVersionNumber(AppName: string): string;
    procedure AssignLicense();
    procedure AddLineToListBox(aListBox: TListBox = nil; aID: Integer = 0; aText: string = '');
  public
    function GetInputFiles(const Filename, Controlname: string): TPWideCharArray;
    function GetOutputFiles(const Filename, Controlname: string): TPWideCharArray;

    function ReadOptions(aOptFile: string; aIsDefOptionsFile: Boolean = False): Boolean;
    procedure RunIteman(data, ctrl, outp, mrf1: string);
    procedure MRFCall(aMRFName: string; aMRFRun: Boolean);
    procedure DialogCall();

    { Thread }
    procedure DoTaskStarted;
    procedure DoTaskFinished;
    procedure DoTaskException;
    procedure StartMyTask(aData: string; aCtrl: string; aOutP: string; aMRF1: string);

    property TreadExceptMessage: string read FTreadExceptMessage write FTreadExceptMessage;
    { }

    property MRFFile: string read FMRFFile write FMRFFile;
    property APPPath: string read FAPPPath write FAPPPath;
    property Is_Demo: Boolean read FIs_Demo write FIs_Demo;
    property MRFText1: string read FMRFText1 write FMRFText1;
    property DocxOutput: string read FDocxOutput write FDocxOutput;
    property UseMRF: Boolean read FUseMRF write FUseMRF;
  end;

var
  frmMain: TfrmMain;

implementation

uses
  FMX.Platform.Win,
  WinAPI.Messages,
  Generics.Defaults;

{$R *.fmx}

var
  // ApplicationDir, OutputDir, outputfilename : TFileName;
  OutputDir: TFileName;
  OldWndProc: TWndProc;
  MinHeight: Integer;

const
  C_DEFAULT_WIDTH = 1316;
  C_DEFAULT_HEIGHT = 807;
  C_MIN_WIDTH = 800;
  C_MIN_HEIGHT = 600;

function NewWndProc(hwnd: hwnd; uMsg: UINT; wParam: wParam; lParam: lParam): LRESULT; stdcall;
var
  PMinMaxInfo: ^TMinMaxInfo;
  TmpSize: TSize;
begin
  if uMsg = WM_GETMINMAXINFO then
  begin
    PMinMaxInfo := Pointer(lParam);
    PMinMaxInfo.ptMinTrackSize.X := C_MIN_WIDTH;
    PMinMaxInfo.ptMinTrackSize.Y := MinHeight;
    Result := 0;
  end
  else
    Result := OldWndProc(hwnd, uMsg, wParam, lParam);
end;

{ TMyTask }

procedure TMyTask.DoStarted();
begin
  if Assigned(OnTaskStarted) then
    OnTaskStarted;
end;

procedure TMyTask.DoFinished();
begin
  if Assigned(OnTaskFinished) then
    OnTaskFinished;
end;

procedure TMyTask.DoEnded();
begin
  if Assigned(OnTaskEnded) then
    OnTaskEnded();
end;

procedure TMyTask.DoException();
begin
  if Assigned(OnTaskException) then
    OnTaskException();
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  ARect: TRect;
  AWinHandle: hwnd;
begin
  FFormatSettings := TFormatSettings.Create;
  FFormatSettings.DecimalSeparator := '.';
  AWinHandle := WindowHandleToPlatform(Handle).Wnd;
  OldWndProc := Pointer(GetWindowLongPtr(AWinHandle, GWLP_WNDPROC));;
  SystemParametersInfo(SPI_GETWORKAREA, 0, @ARect, 0);
  MinHeight := Min(C_MIN_HEIGHT, ARect.Height);
  SetWindowLongPtr(AWinHandle, GWLP_WNDPROC, LongInt(@NewWndProc));

  if (ARect.Width < Width) and (ARect.Height < Height) then
    BorderIcons := [TBorderIcon.biSystemMenu, TBorderIcon.biMinimize];

  if Width > ARect.Width then
    Width := ARect.Width;
  if Height > ARect.Height then
    Height := ARect.Height;
end;

procedure TfrmMain.FormResize(Sender: TObject);
var
  kx: double;
  ky: double;
begin
  kx := Width / C_DEFAULT_WIDTH;
  ky := Height / C_DEFAULT_HEIGHT;
  if kx > 1 then
    kx := 1;
  if ky > 1 then
    ky := 1;
  LayoutContent.Scale.X := kx;
  LayoutContent.Scale.Y := ky;

  Caption := Format('%dx%d', [Width, Height]);
end;

procedure TfrmMain.FormShow(Sender: TObject);
begin
{$REGION ' POPUP WIN '}
  SplashShower();
{$ENDREGION ' POPUP WIN '}
  { Thread }
  pnlLoad.Visible := False;
  Application.ProcessMessages;
  { }

  AddLineToListBox(lbPrecision, 1, '1');
  AddLineToListBox(lbPrecision, 2, '2');
  AddLineToListBox(lbPrecision, 3, '3');

  AddLineToListBox(lbCutPoint, 3, '3');
  AddLineToListBox(lbCutPoint, 4, '4');
  AddLineToListBox(lbCutPoint, 5, '5');
  AddLineToListBox(lbCutPoint, 6, '6');
  AddLineToListBox(lbCutPoint, 7, '7');
  AddLineToListBox(lbCutPoint, 8, '8');
  AddLineToListBox(lbCutPoint, 9, '9');

  btnMenu.Visible := False;
  btnMenu.HitTest := False;

  mtvMenu.Visible := False;

  tbMain.StylesData['Caption.Text'] := 'Iteman';

  tcMain.Tabs[0].IsSelected := True;
  tcMain.TabPosition := TTabPosition.None;

  cbExternalScoreFile.IsChecked := False;
  eExternalScoreFile.Enabled := False;

  eDataMatrixFile.StylesData['Button.Tag'] := eDataMatrixFile.Tag;
  eDataMatrixFile.StylesData['Button.OnClick'] := TValue.From<TNotifyEvent>(btnTab1OnClick);
  eItemControlFile.StylesData['Button.Tag'] := eItemControlFile.Tag;
  eItemControlFile.StylesData['Button.OnClick'] := TValue.From<TNotifyEvent>(btnTab1OnClick);
  eOutputFile.StylesData['Button.Tag'] := eOutputFile.Tag;
  eOutputFile.StylesData['Button.OnClick'] := TValue.From<TNotifyEvent>(btnTab1OnClick);
  eExternalScoreFile.StylesData['Button.Tag'] := eExternalScoreFile.Tag;
  eExternalScoreFile.StylesData['Button.OnClick'] := TValue.From<TNotifyEvent>(btnTab1OnClick);

{$REGION ' Default settings '}
  Is_Demo := False;
  FFirstActivate := True;
  FProtect := TProtect.Create;
  /// /  FExport := TExportIteman.Create;

  APPPath := IncludeTrailingBackslash(ExtractFilePath(ParamStr(0)));
{$ENDREGION ' Default settings '}
  ReadOptions(APPPath + 'Defaults.options', True);

  FProtect.AssignLicense;
  if (FFirstActivate) and not(FProtect.LicenseStatus in [lcsActive, lcsWillExpire]) then
  begin
    FFirstActivate := False;
    Is_Demo := True;

    OpenLicensePad(btnLicense);
  end;

  AssignLicense();
end;

function TfrmMain.GetFloatValue(aFloatValueInStr: string = '0'; aFloatDidgets: Integer = -3): Extended;
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

procedure TfrmMain.btnMenuClick(Sender: TObject);
begin
  mtvMenu.ShowMaster;
end;

procedure TfrmMain.btnMenuItemOnClick(Sender: TObject);
begin
  mtvMenu.HideMaster;

  case TButton(Sender).Tag of
    // Menu File Buttons
    1: // Create AppPath Multiple Runs File
      begin
        ShowMultipleRunsFileForm();
        if Assigned(frmMultipleRunsFile) then
          frmMultipleRunsFile.RunForm(
            procedure()
            begin

            end);
      end;
    2: // Run AppPath Saved Multiple Runs File
      begin
        with OpenDialog do
        begin
          Title := 'Select the multiple run file...';
          InitialDir := APPPath;
          Filename := '';
          Filter := 'TXT files (*.txt)|*.txt|All files (*.*)|*.*';
          FilterIndex := 1;

          if Execute then
            MRFText1 := Filename;
        end;

        if (MRFText1 <> '') then
        begin
          ShowMultipleRunsFileForm();
          frmMultipleRunsFile.RunMRF1(MRFText1, False);
        end;
      end;
    3: // Open an Options File
      begin
        with OpenDialog do
        begin
          Title := 'Select the program options file...';
          InitialDir := APPPath;
          Filename := '';
          Filter := 'OPTIONS files (*.options)|*.options|TXT files (*.txt)|*.txt|All files (*.*)|*.*';
          FilterIndex := 1;

          if Execute then
            if Filename <> '' then
              if ReadOptions(Filename) then
                TDialogService.ShowMessage('Option file successfully uploaded.')
              else
                TDialogService.ShowMessage('Option file do not uploaded.');
        end;
      end;
    4: // Save an Options File
      begin
        with SaveDialog do
        begin
          Title := 'Please enter a file name for the program options file ...';
          InitialDir := APPPath;
          Filename := '';
          Filter := 'OPTIONS files (*.options)|*.options|All files (*.*)|*.*';
          FilterIndex := 1;

          if Execute then
            SaveOptionToFile(Filename);
        end;
      end;
    5: // Save the Program Defaults
      begin
        SaveOptionToFile(APPPath + { '/' + } 'Defaults.options');

        if FileExists('Defaults.options') then
          TDialogService.ShowMessage('The Program Defaults file successfully saved.')
      end;
    // Tabs Menu Buttons
    6: // Files
      begin
        tcMain.Tabs[0].IsSelected := True;
        tcMain.Tabs[1].IsSelected := False;
        tcMain.Tabs[2].IsSelected := False;
        tcMain.Tabs[3].IsSelected := False;

        btnTabFiles.StyleLookup := 'btnMenuItemsActiveStyle';
        btnInputFormat.StyleLookup := 'btnMenuItemsStyle';
        btnScoringOptions.StyleLookup := 'btnMenuItemsStyle';
        btnOutputOptions.StyleLookup := 'btnMenuItemsStyle';
      end;
    7: // Input Format
      begin
        tcMain.Tabs[0].IsSelected := False;
        tcMain.Tabs[1].IsSelected := True;
        tcMain.Tabs[2].IsSelected := False;
        tcMain.Tabs[3].IsSelected := False;

        btnTabFiles.StyleLookup := 'btnMenuItemsStyle';
        btnInputFormat.StyleLookup := 'btnMenuItemsActiveStyle';
        btnScoringOptions.StyleLookup := 'btnMenuItemsStyle';
        btnOutputOptions.StyleLookup := 'btnMenuItemsStyle';
      end;
    8: // Scoring Options
      begin
        tcMain.Tabs[0].IsSelected := False;
        tcMain.Tabs[1].IsSelected := False;
        tcMain.Tabs[2].IsSelected := True;
        tcMain.Tabs[3].IsSelected := False;

        btnTabFiles.StyleLookup := 'btnMenuItemsStyle';
        btnInputFormat.StyleLookup := 'btnMenuItemsStyle';
        btnScoringOptions.StyleLookup := 'btnMenuItemsActiveStyle';
        btnOutputOptions.StyleLookup := 'btnMenuItemsStyle';
      end;
    9: // Output Options
      begin
        tcMain.Tabs[0].IsSelected := False;
        tcMain.Tabs[1].IsSelected := False;
        tcMain.Tabs[2].IsSelected := False;
        tcMain.Tabs[3].IsSelected := True;
        btnInputFormat.StyleLookup := 'btnMenuItemsStyle';
        btnTabFiles.StyleLookup := 'btnMenuItemsStyle';
        btnScoringOptions.StyleLookup := 'btnMenuItemsStyle';
        btnOutputOptions.StyleLookup := 'btnMenuItemsActiveStyle';
      end;
  end;
end;

procedure TfrmMain.SaveOptionToFile(aFileNameWithPath: string = '');
var
  ch, ch1, ch2, ch3, ch4: Char;
  t1, t2, t3: Integer;
  lSaveOpt1: string;
  // lOutputFileOK: Boolean;
  lFile: TextFile;
begin
  lSaveOpt1 := changefileext(aFileNameWithPath, '.options');

  if lSaveOpt1 <> '' then
  begin
    AssignFile(lFile, lSaveOpt1);
    ReWrite(lFile);
    WriteLn(lFile, 'Iteman 4.4');
    WriteLn(lFile, RunTitleBox.text);

    // Line3: Files tab
    if cbDataMatrixFile.IsChecked then
      ch := 'Y'
    else
      ch := 'N';
    Write(lFile, ch, ' ');

    if cbExternalScoreFile.IsChecked then
      ch := 'Y'
    else
      ch := 'N';
    WriteLn(lFile, ch);

    // Line4: Input Format
    t1 := StrToIntDef(eTab2Numbers1.text, 0);
    t2 := StrToIntDef(IDColumnsBegin.text, 0);
    t3 := StrToIntDef(itemColumnBox.text, 0);
    ch := OmitCharBox.text[1];
    ch1 := NACharBox.text[1];
    Write(lFile, t1, ' ', t2, ' ', t3, ' ', ch, ' ', ch1, ' ');

    if DelimitBox.IsChecked then
      ch := 'Y'
    else
      ch := 'N';
    if CommaBox.IsChecked then
      ch1 := 'C'
    else if TabBox.IsChecked then
      ch1 := 'T'
    else
      ch1 := 'N';
    if IncludeIDBox.IsChecked then
      ch2 := 'Y'
    else
      ch2 := 'N';

    WriteLn(lFile, ch, ' ', ch1, ' ', ch2);

    // Line5: DIF options
    if DifCheckBox.IsChecked then
      ch := 'Y'
    else
      ch := 'N'; // DIF related stuff
    Write(lFile, ch, ' ');
    Write(lFile, StrToIntDef(GroupColumn.text, 0), ' ');
    Write(lFile, StrToIntDef(NumDifGroups.text, 0), ' ');
    Write(lFile, Group1Code.text, ' ', Group2Code.text, ' ');
    WriteLn(lFile, Group1label.text, chr(9), Group2label.text);

    // Line6: Scoring options
    if ScaledCheckBox.IsChecked then
      ch := 'Y'
    else
      ch := 'N';
    if ScaledDomain.IsChecked then
      ch1 := 'Y'
    else
      ch1 := 'N';
    if LineARScale.IsChecked then
      ch2 := 'Y'
    else
      ch2 := 'N';

    Write(lFile, ch, ' ', ch1, ' ', ch2, ' ');
    Write(lFile, GetFloatValue(SlopeBox.text, -3), ' ', GetFloatValue(InterceptBox.text, -3), ' ');

    if StandScale.IsChecked then
      ch3 := 'Y'
    else
      ch3 := 'N';
    Write(lFile, ch3, ' ');
    Write(lFile, GetFloatValue(NewMeanBox.text, -3), ' ', GetFloatValue(NewSDBox.text, -3), ' ');

    if ClassCheckBox.IsChecked then
      ch := 'Y'
    else
      ch := 'N';
    if ScoredClass.IsChecked and ClassCheckBox.IsChecked then
      ch1 := 'Y'
    else
      ch1 := 'N';
    if ScaledClass.IsChecked and ClassCheckBox.IsChecked then
      ch2 := 'Y'
    else
      ch2 := 'N';

    Write(lFile, ch, ' ', ch1, ' ', ch2, ' ');
    Write(lFile, GetFloatValue(ClassCutBox.text, -3), ' ');
    WriteLn(lFile, LowLabel.text, chr(9), HighLabel.text);

    // Line7: Output options
    Write(lFile, GetFloatValue(MinPBox.text), ' ', GetFloatValue(MaxPBox.text), ' ');
    Write(lFile, GetFloatValue(MinMeanBox.text), ' ', GetFloatValue(MaxMeanBox.text), ' ');
    Write(lFile, GetFloatValue(MinRpbisBox.text), ' ', GetFloatValue(MaxRpbisBox.text), ' ');

    if OmitAsinc.IsChecked then
      ch := 'Y'
    else
      ch := 'N';
    if SpurBox.IsChecked then
      ch1 := 'Y'
    else
      ch1 := 'N';

    Write(lFile, ch, ' ', ch1, ' ');
    Write(lFile, StrToIntDef(Precision.text, 0), ' ');

    if PlotsBox.IsChecked then
      ch := 'Y'
    else
      ch := 'N';
    if TableBox.IsChecked then
      ch1 := 'Y'
    else
      ch1 := 'N';

    Write(lFile, ch, ' ', ch1, ' ');
    Write(lFile, StrToIntDef(CutPoint.text, 0), ' ');

    if CollusionBox.IsChecked then
      ch := 'Y'
    else
      ch := 'N';
    if ScoredMatrixBox.IsChecked then
      ch1 := 'Y'
    else
      ch1 := 'N';
    if SaveControl.IsChecked then
      ch2 := 'Y'
    else
      ch2 := 'N';
    if IncLomit.IsChecked then
      ch3 := 'Y'
    else
      ch3 := 'N';
    if IncLnr.IsChecked then
      ch4 := 'Y'
    else
      ch4 := 'N';

    WriteLn(lFile, ch, ' ', ch1, ' ', ch2, ' ', ch3, ' ', ch4);

    // Line8: Flags
    Write(lFile, KeyFlag.text, chr(9));
    Write(lFile, LowPFlag.text, chr(9));
    Write(lFile, HighPFlag.text, chr(9));
    Write(lFile, LowRFlag.text, chr(9));
    Write(lFile, HighRFlag.text, chr(9));
    Write(lFile, LowMeanFlag.text, chr(9));
    Write(lFile, HighMeanFlag.text, chr(9));
    WriteLn(lFile, DifFlag.text);

    CloseFile(lFile);
  end;
end;

procedure TfrmMain.btnTab1OnClick(Sender: TObject);
begin
  case TButton(Sender).Tag of
    // Tab 1 buttons click
    1: // Data matrix file
      begin
        with OpenDialog do
        begin
          Title := 'Select the data matrix file...';
          InitialDir := APPPath;
          Filename := '';
          Filter := 'CSV files (*.csv)|*.csv|TXT files (*.txt)|*.txt|DAT files (*.dat)|*.dat|All files (*.*)|*.*';
          FilterIndex := 1;

          if Execute then
            eDataMatrixFile.text := Filename;
        end;
      end;
    2: // Item control file
      begin
        with OpenDialog do
        begin
          Title := 'Select the item control file...';
          InitialDir := APPPath;
          Filename := '';
          Filter := 'CSV files (*.csv)|*.csv|TXT files (*.txt)|*.txt|All files (*.*)|*.*';
          FilterIndex := 1;

          if Execute then
            eItemControlFile.text := Filename;
        end;
      end;
    3: // Output file
      begin
        with SaveDialog do
        begin
          Title := 'Please enter a file name for your output file ...';
          InitialDir := APPPath;
          Filename := '';
          Filter := 'DOCX files (*.docx)|*.docx|All files (*.*)|*.*';
          FilterIndex := 1;

          if Execute then
          begin
            eOutputFile.text := changefileext(Filename, '.docx');
            DocxOutput := eOutputFile.text;
          end;
        end;
      end;
    4: // External score file (optional)
      begin
        with OpenDialog do
        begin
          Title := 'Select the external score file...';
          Filename := '';
          Filter := 'TXT files (*.txt)|*.txt|All files (*.*)|*.*';
          FilterIndex := 1;

          if Execute then
            eExternalScoreFile.text := Filename;
        end;
      end;
  end;
end;

procedure TfrmMain.btnTab2NumbersOnClick(Sender: TObject);
var
  lEdit: TEdit;
  lCurValue: Integer;
  LDValue: double;
  LMaxValue: Integer;
  LMinValue: Integer;
begin
  case TButton(Sender).Tag of
    1, 2:
      lEdit := eTab2Numbers1;
    3, 4:
      lEdit := IDColumnsBegin;
    5, 6:
      lEdit := itemColumnBox;
    7, 8:
      lEdit := GroupColumn;
    9, 10:
      lEdit := NumDifGroups;
    // 11, 12: LEdit := MinPBox;
    // 13, 14: LEdit := MaxPBox;
    // 15, 16: LEdit := MinMeanBox;
    // 17, 18: LEdit := MaxMeanBox;
    // 19, 20: LEdit := MinRpbisBox;
    // 21, 22: LEdit := MaxRpbisBox;
    27, 28:
      lEdit := Precision;
    29, 30:
      lEdit := CutPoint;
  else
    Exit;
  end;

  if TButton(Sender).Tag in [1 .. 10, 15 .. 22, 27 .. 30] then
  begin
    lCurValue := StrToIntDef(lEdit.text, 0);
    case TButton(Sender).Tag of
      // Tab 2 number buttons click
      // Minus
      1, 3, 5, 7, 9, 15, 17, 19, 21, 27, 29:
        inc(lCurValue, -1);
      // Plus
      2, 4, 6, 8, 10, 16, 18, 20, 22, 28, 30:
        inc(lCurValue, 1);
    end;
    lEdit.text := lCurValue.ToString;
  end
  else if TButton(Sender).Tag in [11 .. 14] then
  begin
    LDValue := StrToFloatDef(lEdit.text, 0, FFormatSettings);
    case TButton(Sender).Tag of
      11, 13:
        LDValue := RoundTo(LDValue - 0.01, -2);
      12, 14:
        LDValue := RoundTo(LDValue + 0.01, -2);
    end;
    lEdit.text := FloatToStr(LDValue, FFormatSettings);
  end;
end;

procedure TfrmMain.cbExternalScoreFileChange(Sender: TObject);
begin
  eExternalScoreFile.Enabled := TCheckBox(Sender).IsChecked;
end;

procedure TfrmMain.DelimitBoxChange(Sender: TObject);
begin
  CommaBox.Enabled := TCheckBox(Sender).IsChecked;
  TabBox.Enabled := TCheckBox(Sender).IsChecked;
  IncludeIDBox.Enabled := TCheckBox(Sender).IsChecked;
  IncludeIDBox.IsChecked := TCheckBox(Sender).IsChecked;

  cbDataMatrixFile.IsChecked := False;
  cbDataMatrixFile.Enabled := not TCheckBox(Sender).IsChecked;

  Button1.Enabled := not TCheckBox(Sender).IsChecked;
  eTab2Numbers1.Enabled := not TCheckBox(Sender).IsChecked;
  Button2.Enabled := not TCheckBox(Sender).IsChecked;

  Button3.Enabled := not TCheckBox(Sender).IsChecked;
  IDColumnsBegin.Enabled := not TCheckBox(Sender).IsChecked;
  Button4.Enabled := not TCheckBox(Sender).IsChecked;

  Button5.Enabled := not TCheckBox(Sender).IsChecked;
  itemColumnBox.Enabled := not TCheckBox(Sender).IsChecked;
  Button6.Enabled := not TCheckBox(Sender).IsChecked;

  Text15.Enabled := TCheckBox(Sender).IsChecked;
  Text17.Enabled := TCheckBox(Sender).IsChecked;
  Text18.Enabled := TCheckBox(Sender).IsChecked;

  // Button7.Enabled:= not TCheckBox(Sender).IsChecked;
  // GroupColumn.Enabled:= not TCheckBox(Sender).IsChecked;
  // Button8.Enabled:= not TCheckBox(Sender).IsChecked;
  // Text21.Enabled:= not TCheckBox(Sender).IsChecked;

  IncludeIDBox.Enabled := TCheckBox(Sender).IsChecked;
  IncludeIDBox.IsChecked := TCheckBox(Sender).IsChecked;

  Text16.Enabled := TCheckBox(Sender).IsChecked;
  Text19.Enabled := TCheckBox(Sender).IsChecked;

  TabBox.Enabled := DelimitBox.IsChecked;
  if not TabBox.Enabled then
    TabBox.IsChecked := False;

  CommaBox.IsChecked := DelimitBox.IsChecked;
  if not CommaBox.Enabled then
    CommaBox.IsChecked := False;

  if TCheckBox(Sender).IsChecked then
  begin
    OmitCharBox.Enabled := True;
    NACharBox.Enabled := True;
    GroupColumn.text := '0';
  end;
end;

procedure TfrmMain.DifCheckBoxChange(Sender: TObject);
begin
  Text22.Enabled := TCheckBox(Sender).IsChecked;
  Text23.Enabled := TCheckBox(Sender).IsChecked;
  Text24.Enabled := TCheckBox(Sender).IsChecked;
  Text25.Enabled := TCheckBox(Sender).IsChecked;
  Text26.Enabled := TCheckBox(Sender).IsChecked;

  Button9.Enabled := TCheckBox(Sender).IsChecked;
  Button10.Enabled := TCheckBox(Sender).IsChecked;
  NumDifGroups.Enabled := TCheckBox(Sender).IsChecked;

  Group1Code.Enabled := TCheckBox(Sender).IsChecked;
  Group2Code.Enabled := TCheckBox(Sender).IsChecked;
  Group1label.Enabled := TCheckBox(Sender).IsChecked;
  Group2label.Enabled := TCheckBox(Sender).IsChecked;

  if not DifCheckBox.IsChecked then
  begin
    GroupColumn.text := '0';
    GroupColumn.Enabled := False;
    Button7.Enabled := TCheckBox(Sender).IsChecked;
    Button8.Enabled := TCheckBox(Sender).IsChecked;
    Text21.Enabled := TCheckBox(Sender).IsChecked;
  end
  else if not DelimitBox.IsChecked then
  begin
    GroupColumn.Enabled := True;
    Button7.Enabled := TCheckBox(Sender).IsChecked;
    Button8.Enabled := TCheckBox(Sender).IsChecked;
    Text21.Enabled := TCheckBox(Sender).IsChecked;
  end;
end;

procedure TfrmMain.ScaledCheckBoxChange(Sender: TObject);
begin
  if TCheckBox(Sender).IsChecked then
  begin
    LineARScale.Enabled := TCheckBox(Sender).IsChecked;
    StandScale.Enabled := TCheckBox(Sender).IsChecked;
    Text30.Enabled := TCheckBox(Sender).IsChecked;
    Text32.Enabled := TCheckBox(Sender).IsChecked;
    SlopeBox.Enabled := TCheckBox(Sender).IsChecked;
    NewMeanBox.Enabled := TCheckBox(Sender).IsChecked;
    Text31.Enabled := TCheckBox(Sender).IsChecked;
    Text33.Enabled := TCheckBox(Sender).IsChecked;
    InterceptBox.Enabled := TCheckBox(Sender).IsChecked;
    NewSDBox.Enabled := TCheckBox(Sender).IsChecked;
  end
  else if not(ScaledCheckBox.IsChecked) and not(ScaledDomain.IsChecked) then
  begin
    LineARScale.Enabled := TCheckBox(Sender).IsChecked;
    StandScale.Enabled := TCheckBox(Sender).IsChecked;
    Text30.Enabled := TCheckBox(Sender).IsChecked;
    Text32.Enabled := TCheckBox(Sender).IsChecked;
    SlopeBox.Enabled := TCheckBox(Sender).IsChecked;
    NewMeanBox.Enabled := TCheckBox(Sender).IsChecked;
    Text31.Enabled := TCheckBox(Sender).IsChecked;
    Text33.Enabled := TCheckBox(Sender).IsChecked;
    InterceptBox.Enabled := TCheckBox(Sender).IsChecked;
    NewSDBox.Enabled := TCheckBox(Sender).IsChecked;
  end;
  if not LineARScale.Enabled then
    LineARScale.IsChecked := False;
  if not StandScale.Enabled then
    StandScale.IsChecked := False;
end;

procedure TfrmMain.CheckOptionValue(const AEdit: TEdit);
var
  LMin: variant;
  LMax: variant;
begin
  case AEdit.Tag of
    1:
      begin
        LMin := 0;
        LMax := 1;
        AEdit.text := FloatToStr(RoundTo(SetValueInRange<double>(StrToFloatDef(AEdit.text, 0, FFormatSettings), LMin,
          LMax), -2), FFormatSettings);
      end;
    2:
      begin
        LMin := 1;
        LMax := 4;
        AEdit.text := IntToStr(SetValueInRange<Integer>(StrToIntDef(AEdit.text, 0), LMin, LMax));
      end;
    3:
      begin
        LMin := 3;
        LMax := 7;
        AEdit.text := IntToStr(SetValueInRange<Integer>(StrToIntDef(AEdit.text, 0), LMin, LMax));
      end;
  else
    begin
      LMin := 0;
      LMax := 99999;
      AEdit.text := IntToStr(SetValueInRange<Integer>(StrToIntDef(AEdit.text, 0), LMin, LMax));
    end;
  end;
end;

procedure TfrmMain.ClassCheckBoxChange(Sender: TObject);
begin
  ScoredClass.Enabled := TCheckBox(Sender).IsChecked;
  ScaledClass.Enabled := TCheckBox(Sender).IsChecked;

  if not TCheckBox(Sender).IsChecked then
  begin
    ScoredClass.IsChecked := False;
    ScaledClass.IsChecked := False;
  end
  else
    ScoredClass.IsChecked := not ScaledClass.IsChecked;

  Text38.Enabled := TCheckBox(Sender).IsChecked;
  ClassCutBox.Enabled := TCheckBox(Sender).IsChecked;
  Text40.Enabled := TCheckBox(Sender).IsChecked;
  LowLabel.Enabled := TCheckBox(Sender).IsChecked;
  Text41.Enabled := TCheckBox(Sender).IsChecked;
  HighLabel.Enabled := TCheckBox(Sender).IsChecked;
end;

procedure TfrmMain.ScoredClassClick(Sender: TObject);
begin
  if (Sender = ScoredClass) and TCheckBox(Sender).IsChecked then
  begin
    // ClassCutBox.digits:= 0;
    // ClassCutBox.MinValue:= 0;
    ClassCutBox.hint := 'Range is 0 to 10000';
  end
  else if (Sender = ScaledClass) and TCheckBox(Sender).IsChecked then
  begin
    // ClassCutBox.digits:= 3;
    // ClassCutBox.MinValue:= -10;
    ClassCutBox.hint := 'Range is -10 to 10000';
  end;
end;

procedure TfrmMain.ScoredMatrixBoxChange(Sender: TObject);
begin
  IncLomit.Enabled := TCheckBox(Sender).IsChecked;
  IncLnr.Enabled := TCheckBox(Sender).IsChecked;
end;

function TfrmMain.SetValueInRange<T>(AValue, AMin, AMax: T): T;
var
  Comparer: IComparer<T>;
begin
  Comparer := TComparer<T>.Default;
  if Comparer.Compare(AValue, AMax) > 0 then
    Result := AMax
  else if Comparer.Compare(AValue, AMin) < 0 then
    Result := AMin
  else
    Result := AValue;
end;

{$REGION ' POPUP WIN '}

procedure TfrmMain.OpenLicensePad(Sender: TObject);
begin
  mtvMenu.HideMaster;

  ShowLicenseForm();
  frmLicense.LayoutContent.Parent := pSplashContent;
  frmLicense.btnRequestLicenseKey.OnClick := btnSplashClick;
  frmLicense.btnConfirmLicense.OnClick := btnSplashClick;
  frmLicense.btnBack.OnClick := btnSplashClick;

  SplashShower(True, 400, 640);
end;

procedure TfrmMain.PanelSplashBGClick(Sender: TObject);
begin
  btnSplashClick(nil);
end;

procedure TfrmMain.SplashShower(aMode: Boolean = False; aDefaultHeight: Integer = 300; aSplashWidth: Integer = 300;
aPanelBGStyleName: string = 'PanelSplashGrayStyle');
begin
  if aMode then
  begin
    LayoutSplash.Height := 0;
    LayoutSplashContainer.Width := aSplashWidth;
    LayoutSplashContainer.Height := aDefaultHeight;
  end;

  LayoutSplash.Align := TAlignLayout.Top;
  LayoutSplash.Width := Self.Width;

  FloatAnimation1.Inverse := not aMode;
  FloatAnimation1.StopValue := Self.Height;
  LayoutSplash.Visible := aMode;

  PanelSplashBG.StyleLookup := aPanelBGStyleName;
  PanelSplashBG.HitTest := aMode;
  pSplashContent.HitTest := aMode;

  FloatAnimation1.Start;
end;

procedure TfrmMain.btnSplashClick(Sender: TObject);
begin
  if not Assigned(Sender) then // Cancel
    SplashShower()
  else
  begin
    if Assigned(frmLicense) then
    begin
      case TButton(Sender).Tag of
        0: // btnBack
          begin

          end;
        1: // btnRequestLicenseKey
          begin
            RequestLicenseKey();
          end;
        2: // btnConfirmLicense
          begin
            ConfirmLicense();
          end;
      else
        TDialogService.ShowMessage('No functionality assigned');
      end;

      frmLicense.Close;
      AssignLicense();
    end;

    SplashShower();
  end;
end;
{$ENDREGION ' POPUP WIN '}

procedure TfrmMain.AssignLicense();
begin
  FProtect.Free;
  FProtect := TProtect.Create;
  FProtect.AssignLicense;

  UpdateLicenseDesc();
end;

procedure TfrmMain.RequestLicenseKey();
var
  lType: string;
  lBody: string;
  lCommand: string;
begin
  if Assigned(frmLicense) then
    with frmLicense do
      if (rbLicType365.IsChecked) or (rbLicType180.IsChecked) then
      begin
        if (rbLicType365.IsChecked) then
          lType := '365-DAYS'
        else
          lType := '180-DAYS';

        lBody := '&body=' + '%0D%0A' + '%0D%0A### DO NOT MODIFY OR DELETE THE FOLLOWING LINES ###' + '%0D%0A### APPID: '
          + Protect.Key + ' ###' + '%0D%0A### LIC TYPE: ' + lType + ' ###' + '%0D%0A### END ###';

        lCommand := 'mailto:activations@assess.com';
        lCommand := lCommand + '?subject=Iteman License Key Request' + lBody;

        ShellExecute(0, 'open', PWideChar(lCommand), nil, nil, 1);
      end
      else
        TDialogService.ShowMessage('Pick a license type.');
end;

procedure TfrmMain.ConfirmLicense();
begin
  if Assigned(frmLicense) then
    with frmLicense do
    begin
      Protect.License := eLicenseKey.text;
      try
        Protect.AssignLicense;
      except
        Protect.License := '';
      end;
      if (Protect.LicenseStatus in [lcsActive, lcsWillExpire]) then
        TDialogService.ShowMessage('Valid License Key.' + #13#10 + 'Your Iteman application is activated.')
      else
        TDialogService.ShowMessage('Invalid License Key!');
    end;
end;

function TfrmMain.ReadOptions(aOptFile: string; aIsDefOptionsFile: Boolean = False): Boolean;
var
  lVerName: string;
  lName1: string;
  lName2: string;
  ch, ch1, ch2, ch3, ch4: Char;
  t1, t2, t3, i, j, k: Integer;
  r1, r2, r3, r4, r5, r6: Real;
  labels2: array [1 .. 2] of string;

  lFile: TextFile;
begin
  Result := True;

  if aIsDefOptionsFile then
    lName2 := 'The Defaults.options file could not be found or it does not begin with "Iteman 4.4" ' + #13 +
      'or is otherwise invalid. You can create a new options file after completing the ' + #13 +
      'options on all tabs then saving the file using "Save the Program Defaults" on the ' + #13 +
      'File pull-down menu.'
  else
    lName2 := 'The options file is invalid or it does not begin with "Iteman 4.4". ' + #13 +
      'Please create a new options file using "Save an Options File" on the' + #13 +
      'File pull-down menu and try again.';

  AssignFile(lFile, aOptFile);
  Reset(lFile);
  ReadLn(lFile, lVerName);
  ReadLn(lFile, lName1);
  RunTitleBox.text := lName1;

  ch := chr(0);
  ch1 := chr(0);
  // Line 3 Files tab
  try
    ReadLn(lFile, ch, ch1, ch1);
  except
    TDialogService.ShowMessage(lName2);
  end;

  cbDataMatrixFile.IsChecked := (UpperCase(ch) = 'Y');
  cbExternalScoreFile.IsChecked := (UpperCase(ch1) = 'Y');

  ch1 := chr(0);
  ch2 := chr(0);
  ch3 := chr(0);
  ch4 := chr(0);

  // Line 4 Input Format tab
  try
    ReadLn(lFile, t1, t2, t3, ch, ch, ch1, ch1, ch2, ch2, ch3, ch3, ch4, ch4);
  except
    Result := False;
  end;

  eTab2Numbers1.text := t1.ToString;
  IDColumnsBegin.text := t2.ToString;
  itemColumnBox.text := t3.ToString;
  OmitCharBox.text := ch;
  NACharBox.text := ch1;

  DelimitBox.IsChecked := (UpCase(ch2) = 'Y');
  DelimitBox.OnChange(DelimitBox);
  if DelimitBox.IsChecked then
  begin
    if UpperCase(ch3) = 'C' then
      CommaBox.IsChecked := True
    else if UpperCase(ch3) = 'T' then
      TabBox.IsChecked := True;
  end;

  IncludeIDBox.IsChecked := (UpCase(ch4) = 'Y');

  r1 := 0;
  r2 := 0;
  // Line 5 DIF options
  try
    Read(lFile, ch, r1, r2, ch1, ch1, ch2, ch2, ch3);
  except
    Result := False;
  end;

  if (ch = chr(26)) or (ch1 = chr(26)) or (ch2 = chr(26)) or (ch3 = chr(26)) then
    Result := False;

  if Result then
  begin
    DifCheckBox.IsChecked := (UpperCase(ch) = 'Y');

    GroupColumn.text := FloatToStr(r1, FFormatSettings);
    NumDifGroups.text := FloatToStr(r2, FFormatSettings);
    Group1Code.text := UpperCase(ch1);
    Group2Code.text := UpperCase(ch2);

    for j := 1 to 2 do
    begin
      i := 0;
      ch := chr(0);

      while (ch <> chr(9)) and (ch <> chr(13)) do
      begin
        inc(i);
        Read(lFile, ch);
        ReadFlags[i] := ch;
      end;

      ReadFlags[i] := chr(0); // Reset last character (space) to null
      labels2[j] := pchar(@ReadFlags);

      for k := 1 to i do
        ReadFlags[k] := chr(0);
    end;

    if ch <> chr(13) then
      ReadLn(lFile);
    if ch = chr(13) then
      read(lFile, ch); // end line

    Group1label.text := labels2[1];
    Group2label.text := labels2[2];
  end;

  r3 := 0;
  r4 := 0;
  // Line 6 Scoring options
  try
    Read(lFile, ch, ch1, ch1, ch2, ch2, r1, r2, ch3, ch3, r3, r4);
  except
    Result := False;
  end;

  ScaledCheckBox.IsChecked := (UpperCase(ch) = 'Y');
  ScaledDomain.IsChecked := (UpperCase(ch1) = 'Y');
  ScaledCheckBox.OnChange(ScaledCheckBox);
  if ScaledCheckBox.IsChecked then
  begin
    LineARScale.IsChecked := (UpperCase(ch2) = 'Y');
    if LineARScale.IsChecked then
    begin
      SlopeBox.text := FloatToStr(r1);
      InterceptBox.text := FloatToStr(r2);
    end;
    StandScale.IsChecked := (UpperCase(ch3) = 'Y');
    if StandScale.IsChecked then
    begin
      NewMeanBox.text := FloatToStr(r3);
      NewSDBox.text := FloatToStr(r4);
    end;
  end
  else
    StandScaleChange(StandScale);

  try
    Read(lFile, ch, ch, ch1, ch1, ch2, ch2, r1, ch3);
  except
    Result := False;
  end;

  ClassCheckBox.IsChecked := (UpperCase(ch) = 'Y');
  if ClassCheckBox.IsChecked then
  begin
    ScoredClass.IsChecked := (UpperCase(ch1) = 'Y');
    ScaledClass.IsChecked := (UpperCase(ch2) = 'Y');
    ScoredClass.IsChecked := not ScaledClass.IsChecked;
  end
  else
  begin
    ScoredClass.IsChecked := False;
    ScaledClass.IsChecked := False;
  end;
  ClassCheckBoxChange(ClassCheckBox);

  ClassCutBox.text := FloatToStr(r1);

  if Result then
  begin
    for j := 1 to 2 do
    begin
      i := 0;
      ch := chr(0);

      while (ch <> chr(9)) and (ch <> chr(13)) do
      begin
        inc(i);
        Read(lFile, ch);
        ReadFlags[i] := ch;
      end;

      ReadFlags[i] := chr(0); // Reset last character (space) to null
      labels2[j] := PWideChar(@ReadFlags);
      for k := 1 to i do
        ReadFlags[k] := chr(0);
    end;

    if ch <> chr(13) then
      ReadLn(lFile);
    if ch = chr(13) then
      read(lFile, ch); // end line

    LowLabel.text := labels2[1];
    HighLabel.text := labels2[2];
  end;

  r1 := 0;
  r2 := 0;
  r3 := 0;
  r4 := 0;
  r5 := 0;
  r6 := 0;

  // Line 7 Output Options
  try
    Read(lFile, r1, r2, r3, r4, r5, r6, ch, ch, ch1, ch1, t1, ch2, ch2, ch3, ch3);
  except
    Result := False;
  end;

  MinPBox.text := FloatToStr(r1, FFormatSettings);
  MaxPBox.text := FloatToStr(r2, FFormatSettings);
  MinMeanBox.text := FloatToStr(r3, FFormatSettings);
  MaxMeanBox.text := FloatToStr(r4, FFormatSettings);
  MinRpbisBox.text := FloatToStr(r5, FFormatSettings);
  MaxRpbisBox.text := FloatToStr(r6, FFormatSettings);
  OmitAsinc.IsChecked := (UpperCase(ch) = 'Y');
  SpurBox.IsChecked := (UpperCase(ch1) = 'Y');
  Precision.text := t1.ToString;
  PlotsBox.IsChecked := (UpperCase(ch2) = 'Y');
  TableBox.IsChecked := (UpperCase(ch3) = 'Y');

  try
    ReadLn(lFile, t1, ch, ch, ch1, ch1, ch2, ch2, ch3, ch3, ch4, ch4);
  except
    Result := False;
  end;

  CutPoint.text := t1.ToString;
  CollusionBox.IsChecked := (UpperCase(ch) = 'Y');
  ScoredMatrixBox.IsChecked := (UpperCase(ch1) = 'Y');
  SaveControl.IsChecked := (UpperCase(ch2) = 'Y');
  IncLomit.IsChecked := (UpperCase(ch3) = 'Y');
  IncLnr.IsChecked := (UpperCase(ch4) = 'Y');

  if (UpperCase(ch1) <> 'Y') then
  begin
    IncLomit.IsChecked := False;
    IncLnr.IsChecked := False;
  end;

  // Line 8 Flags
  for j := 1 to 8 do
  begin
    i := 0;
    ch := chr(0);

    while (ch <> chr(9)) and (ch <> chr(13)) do
    begin
      inc(i);

      if ch = chr(26) then
      begin
        Result := False;
        break;
      end;

      Read(lFile, ch);
      ReadFlags[i] := ch;
    end;

    ReadFlags[i] := chr(0); // Reset last character (space) to null
    FlagArray[j] := PWideChar(@ReadFlags);

    for k := 1 to i do
      ReadFlags[k] := chr(0);

    if not Result then
      break;
  end;

  KeyFlag.text := FlagArray[1];
  LowPFlag.text := FlagArray[2];
  HighPFlag.text := FlagArray[3];
  LowRFlag.text := FlagArray[4];
  HighRFlag.text := FlagArray[5];
  LowMeanFlag.text := FlagArray[6];
  HighMeanFlag.text := FlagArray[7];
  DifFlag.text := FlagArray[8];

  if Result = False then
    TDialogService.ShowMessage(lName2);

  CloseFile(lFile);
end;

procedure TfrmMain.MRFCall(aMRFName: string; aMRFRun: Boolean);
begin
  if aMRFRun then
    TDialogService.ShowMessage('ITEMAN has completed the multiple run(s) analysis.')
  else
    TDialogService.ShowMessage
      ('The multiple runs analysis could not be completed. Please check the file specifications.');
end;

procedure TfrmMain.UpdateLicenseDesc();
var
  lDTExp, lDTNow: TDate;
  lTxt: string;
  lFmt: TFormatSettings;
begin
  if FProtect.LicenseStatus in [lcsDemo, lcsExpired] then
  begin
    tLicenceDaysLeft.text := 'demo';
    tLicenceDaysLeft.TextSettings.FontColor := $FFFC0606;
    tLicenceExpiresDate.text := 'demo';
    tLicenceExpiresDate.TextSettings.FontColor := $FFFC0606;
  end
  else
  begin
    FProtect.DecodeLicense(lTxt, lDTExp);
    lDTNow := Now;
    tLicenceDaysLeft.text := IntToStr(DaysBetween(lDTExp, lDTNow));
    // < add +1 if Human perception
    tLicenceDaysLeft.TextSettings.FontColor := $FF272727;
    lFmt := FormatSettings;
    lFmt.DateSeparator := '/';
    tLicenceExpiresDate.text := FormatDateTime('dd/mm/yyyy', lDTExp, lFmt);
    tLicenceExpiresDate.TextSettings.FontColor := $FF272727;
  end;
end;

procedure TfrmMain.btnRunClick(Sender: TObject);
begin
  RunIteman(eDataMatrixFile.text, eItemControlFile.text, eOutputFile.text, 'GUI');
end;

{$REGION ' RunIteman '}

procedure TfrmMain.RunIteman(data, ctrl, outp, mrf1: string);
begin
  StartMyTask(data, ctrl, outp, mrf1);
end;

{ Thread }
procedure TMyTask.Execute();
var
  lItemanCalculation: TItemanCalculation;
begin
  Synchronize(DoStarted);
  try
{$REGION ' My tasks '}
    try
      lItemanCalculation := TItemanCalculation.Create(frmMain.FlagArray);
      lItemanCalculation.data := InputData;
      lItemanCalculation.ctrl := InputCtrl;
      lItemanCalculation.outp := InputOutp;
      lItemanCalculation.mrf1 := InputMrf1;

      if not lItemanCalculation.RunIteman(StrToIntDef(frmMain.itemColumnBox.text, 0),
        GetDataMatrixVariableAdditionalColumns, ItemControlVariableAdditionalRows) then
      begin
        frmMain.TreadExceptMessage := lItemanCalculation.ErrorMessage;
        frmMain.UseMRF := lItemanCalculation.UseMRF;
        Synchronize(DoException);
        Exit;
      end
      else
      begin
        frmMain.TreadExceptMessage := lItemanCalculation.ErrorMessage;
        frmMain.UseMRF := lItemanCalculation.UseMRF;
        Synchronize(DoException);
      end;
    except
      on E: Exception do
      begin
        frmMain.TreadExceptMessage := Format('Internal error:' + sLineBreak + '%s', [E.ToString]);
        Synchronize(DoException);
        Exit;
      end;
    end;
  finally
    Synchronize(DoFinished);
  end;
end;

function TMyTask.GetDataMatrixVariableAdditionalColumns: Integer;
begin
  if frmMain.IncludeIDBox.IsChecked then
    Result := 1
  else
    Result := 0;
end;

procedure TfrmMain.DoTaskStarted;
begin
  pnlLoad.Visible := True;
  AniIndicator1.Enabled := True;
  Application.ProcessMessages;
end;

procedure TfrmMain.eTab2Numbers1Change(Sender: TObject);
var
  lEdit: TEdit;
begin
  if Sender is TEdit then
    CheckOptionValue(TEdit(Sender));
end;

procedure TfrmMain.DoTaskFinished;
begin
  if Assigned(AniIndicator1) then
  begin
    AniIndicator1.Enabled := False;
    pnlLoad.Visible := False;
    Application.ProcessMessages;
    LayoutContent.Repaint;

    if not UseMRF then
      DialogCall();
  end;
end;

procedure TfrmMain.DoTaskException;
begin
  AniIndicator1.Enabled := False;
  pnlLoad.Visible := False;
  Application.ProcessMessages;
  LayoutContent.Repaint;

  TDialogService.ShowMessage(TreadExceptMessage);
end;

procedure TfrmMain.StandScaleChange(Sender: TObject);
begin
  if Sender = StandScale then
  begin
    NewMeanBox.Enabled := StandScale.IsChecked;
    if not NewMeanBox.Enabled then
      NewMeanBox.text := '0.000';
    NewSDBox.Enabled := StandScale.IsChecked;
    if not NewSDBox.Enabled then
      NewSDBox.text := '0.000';
  end
  else if Sender = LineARScale then
  begin
    SlopeBox.Enabled := LineARScale.IsChecked;
    if not SlopeBox.Enabled then
      SlopeBox.text := '0.000';
    InterceptBox.Enabled := LineARScale.IsChecked;
    if not InterceptBox.Enabled then
      InterceptBox.text := '0.000';
  end;
end;

procedure TfrmMain.StartMyTask(aData: string; aCtrl: string; aOutP: string; aMRF1: string);
begin
  if Assigned(FMyTask) then
    with FMyTask do
    begin
      OnTaskStarted := nil;
      OnTaskFinished := nil;
      OnTaskEnded := nil;
      OnTaskException := nil;
      Free;
    end;

  FMyTask := TMyTask.Create(True);
  with FMyTask do
  begin
    OnTaskStarted := DoTaskStarted;
    OnTaskFinished := DoTaskFinished;
    OnTaskEnded := nil;
    OnTaskException := DoTaskException;

    { Parameters }
    InputData := aData;
    InputCtrl := aCtrl;
    InputOutp := aOutP;
    InputMrf1 := aMRF1;

    Start;
  end;
end;
{ }

procedure TfrmMain.DialogCall();
begin
  TDialogService.MessageDialog('The run is complete.' + #13 +
    'The summary output file can be found at the following Windows path:' + #13 + DocxOutput + #13 + #13 +
    'Do you want to open the directory for the output file?' + #13, TMsgDlgType.mtConfirmation,
    [TMsgDlgBtn.mbYes, TMsgDlgBtn.mbNo], TMsgDlgBtn.mbYes, 0,
    procedure(const AResult: TModalResult)
    begin
      if AResult = mrYes then
      begin
        FInst := ShellExecute(0, PWideChar('open'), PWideChar(ExtractFilePath(DocxOutput)), PWideChar(''),
          PWideChar(''), SW_SHOWNORMAL);

        if (FInst <= 32) then
          TDialogService.ShowMessage('Unable to open the directory to the output file.');
      end;
    end);
end;

procedure TfrmMain.GetBuildInfo(const SourceFile: String; var V1, V2, V3, V4: Word);
var
  VerInfoSize: DWORD;
  VerInfo: Pointer;
  VerValueSize: DWORD;
  VerValue: PVSFixedFileInfo;
  Dummy: DWORD;
begin
  VerInfoSize := GetFileVersionInfoSize(pchar(SourceFile), Dummy);

  if VerInfoSize = 0 then
    Exit;

  GetMem(VerInfo, VerInfoSize);
  GetFileVersionInfo(pchar(SourceFile), 0, VerInfoSize, VerInfo);
  VerQueryValue(VerInfo, '\', Pointer(VerValue), VerValueSize);

  with VerValue^ do
  begin
    V1 := dwFileVersionMS shr 16;
    V2 := dwFileVersionMS and $FFFF;
    V3 := dwFileVersionLS shr 16;
    V4 := dwFileVersionLS and $FFFF;
  end;

  FreeMem(VerInfo, VerInfoSize);
end;

function TfrmMain.GetVersionNumber(AppName: string): string;
var
  V1, V2, V3, V4: Word;
begin
  GetBuildInfo(AppName, V1, V2, V3, V4);
  Result := IntToStr(V1) + '.' + IntToStr(V2) + '.' + IntToStr(V3) + '.' + IntToStr(V4);
end;

function TfrmMain.GetInputFiles(const Filename, Controlname: string): TPWideCharArray;
var
  ifor: Integer;
  List: TStringList;
begin
  List := TStringList.Create;
  try
    if UseMRF then
    begin
      List.Add(Filename);

      if not cbDataMatrixFile.IsChecked then
        List.Add(Controlname);
    end
    else
    begin
      List.Add(eDataMatrixFile.text);

      if not cbDataMatrixFile.IsChecked then
        List.Add(eItemControlFile.text);

      if cbExternalScoreFile.IsChecked then
        List.Add(eExternalScoreFile.text);
    end;

    SetLength(Result, List.Count);

    for ifor := 0 to pred(List.Count) do
      Result[ifor] := PWideChar(List[ifor]);
  finally
    List.Free;
  end;
end;

function TfrmMain.GetOutputFiles(const Filename, Controlname: string): TPWideCharArray;
var
  ifor: Integer;
  List: TStringList;
begin
  List := TStringList.Create;
  try
    if UseMRF then
    begin
      List.Add(changefileext(Filename, '.docx'));
      List.Add(changefileext(Filename, '.csv'));
      List.Add(changefileext(Filename, '') + ' Scores' + '.csv');

      if ScoredMatrixBox.IsChecked then
        List.Add(changefileext(Filename, '') + ' Matrix' + '.txt');
    end
    else
    begin
      List.Add(eOutputFile.text);
      List.Add(changefileext(eOutputFile.text, '.csv'));
      List.Add(changefileext(changefileext(eOutputFile.text, '') + ' Scores', '.csv'));
      List.Add(changefileext(changefileext(eOutputFile.text, '') + ' Scores', '.csv'));

      if ScoredMatrixBox.IsChecked then
        List.Add(changefileext(changefileext(eOutputFile.text, '') + ' Matrix', '.txt'));

      OutputDir := ExtractFilePath(eOutputFile.text);
    end;

    SetLength(Result, List.Count);

    for ifor := 0 to pred(List.Count) do
      Result[ifor] := PWideChar(List[ifor]);
  finally
    List.Free;
  end;
end;

procedure TfrmMain.AddLineToListBox(aListBox: TListBox = nil; aID: Integer = 0; aText: string = '');
var
  lLBItem: TListBoxItem;
begin
  lLBItem := TListBoxItem.Create(nil);

  lLBItem.StyleLookup := 'ComboBoxItemStyle';
  lLBItem.Tag := aID;
  lLBItem.text := aText;

  lLBItem.StylesData['background.Fill.Color'] := TValue.From<TAlphaColor>($01010101);

  aListBox.AddObject(lLBItem);
end;

procedure TfrmMain.lbPrecisionItemClick(const Sender: TCustomListBox; const Item: TListBoxItem);
var
  ifor: Integer;
begin
  if (Item.Tag <> 0) then
  begin
    Precision.text := Item.text;

    for ifor := 0 to pred(TListBox(Sender).Count) do
      TListBox(Sender).ItemByIndex(ifor).IsSelected := (TListBox(Sender).ItemByIndex(ifor).Tag = Item.Tag);

    popPrecisionList.IsOpen := False;
  end;
end;

procedure TfrmMain.lbCutPointItemClick(const Sender: TCustomListBox; const Item: TListBoxItem);
var
  ifor: Integer;
begin
  if (Item.Tag <> 0) then
  begin
    CutPoint.text := Item.text;

    for ifor := 0 to pred(TListBox(Sender).Count) do
      TListBox(Sender).ItemByIndex(ifor).IsSelected := (TListBox(Sender).ItemByIndex(ifor).Tag = Item.Tag);

    popCutPointList.IsOpen := False;
  end;
end;

end.
