unit ufrmMultipleRunsFile;

interface

uses
  System.SysUtils,
  System.Types,
  System.UITypes,
  System.Classes,
  System.Variants,

  System.Rtti,
  System.StrUtils,
  vkbdhelper,
  FMX.DialogService,

  FMX.Types,
  FMX.Controls,
  FMX.Forms,
  FMX.Graphics,
  FMX.Dialogs,
  FMX.Objects,
  FMX.Layouts,
  FMX.StdCtrls,
  FMX.Controls.Presentation,
  FMX.Edit,
  FMX.ScrollBox,
  FMX.Memo,
  FMX.Memo.Types;

type
  TfrmMultipleRunsFile = class(TForm)
    LayoutContent: TLayout;
    tbMain: TToolBar;
    btnBack: TButton;
    gplMain: TGridPanelLayout;
    Layout12: TLayout;
    gplTopInfoBlock: TGridPanelLayout;
    Layout13: TLayout;
    Layout18: TLayout;
    gplBottom: TGridPanelLayout;
    Layout42: TLayout;
    btnRun: TButton;
    Layout43: TLayout;
    btnRunMRF: TButton;
    gplDataContent: TGridPanelLayout;
    Layout1: TLayout;
    gplTab1Row1: TGridPanelLayout;
    Layout26: TLayout;
    Text2: TText;
    Layout27: TLayout;
    pathedit: TEdit;
    Layout2: TLayout;
    addpath: TButton;
    Layout3: TLayout;
    GridPanelLayout1: TGridPanelLayout;
    Layout4: TLayout;
    Text1: TText;
    Layout5: TLayout;
    OptionEdit: TEdit;
    Layout6: TLayout;
    AddOption: TButton;
    Layout7: TLayout;
    UseDefault: TButton;
    Layout8: TLayout;
    GridPanelLayout2: TGridPanelLayout;
    Layout9: TLayout;
    Text3: TText;
    Layout10: TLayout;
    controledit: TEdit;
    Layout11: TLayout;
    AddControl: TButton;
    Layout14: TLayout;
    SkipControl: TButton;
    Layout15: TLayout;
    GridPanelLayout3: TGridPanelLayout;
    Layout16: TLayout;
    Text4: TText;
    Layout17: TLayout;
    dataedit: TEdit;
    Layout19: TLayout;
    AddData: TButton;
    Layout20: TLayout;
    Layout21: TLayout;
    currentMRF: TMemo;
    SaveDialog: TSaveDialog;
    OpenDialog: TOpenDialog;
    SelectDirectoryDialog: TOpenDialog;
    GoodLogo: TImage;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnPathButtonOnClick(Sender: TObject);
    procedure btnOptionsOnClick(Sender: TObject);
    procedure btnBackClick(Sender: TObject);
    procedure btnRunClick(Sender: TObject);
    procedure btnRunMRFClick(Sender: TObject);
  private
    FSuccProc: TProc;
    FFormResult: TModalResult;

    FDATA: string;
  public
    procedure RunForm(const aSuccProc: TProc);

    procedure RunMRF1(aMRFFileName: string = ''; aSaveMRF: Boolean = False);

    property FormResult: TModalResult read FFormResult write FFormResult;
  end;

function ShowMultipleRunsFileForm(): TfrmMultipleRunsFile;

var
  frmMultipleRunsFile: TfrmMultipleRunsFile;

  LineNumber: Integer;
  MRFFile, Path, OPTS, CTRL: string;
  MRFText: string;
  MRFRun, MRFOk, PathOk, OPTSOk, CTRLOk, DataOk: Boolean;

implementation

uses
  ufrmMain;

{$R *.fmx}

function ShowMultipleRunsFileForm(): TfrmMultipleRunsFile;
begin
  frmMultipleRunsFile := TfrmMultipleRunsFile.Create(Application);

  with frmMultipleRunsFile do
  begin
    tbMain.StylesData['Caption.Text'] := 'Multiple Runs Creator';

    pathedit.StylesData['Button.Tag'] := pathedit.Tag;
    pathedit.StylesData['Button.OnClick'] := TValue.From<TNotifyEvent>(btnPathButtonOnClick);
    OptionEdit.StylesData['Button.Tag'] := OptionEdit.Tag;
    OptionEdit.StylesData['Button.OnClick'] := TValue.From<TNotifyEvent>(btnPathButtonOnClick);
    controledit.StylesData['Button.Tag'] := controledit.Tag;
    controledit.StylesData['Button.OnClick'] := TValue.From<TNotifyEvent>(btnPathButtonOnClick);
    dataedit.StylesData['Button.Tag'] := dataedit.Tag;
    dataedit.StylesData['Button.OnClick'] := TValue.From<TNotifyEvent>(btnPathButtonOnClick);

    AddOption.Enabled := False;
    UseDefault.Enabled := False;
    AddControl.Enabled := False;
    SkipControl.Enabled := False;
    AddData.Enabled := False;
    btnRunMRF.Enabled := False;

    LineNumber := 0;
  end;

  Result := frmMultipleRunsFile;
end;

procedure TfrmMultipleRunsFile.RunForm(const aSuccProc: TProc);
begin
  FSuccProc := aSuccProc;
  FormResult := mrCancel;

{$IF DEFINED(Win64) or DEFINED(Win32)}
  ShowModal;
{$ELSE}
  Self.Show;
{$ENDIF}
end;

procedure TfrmMultipleRunsFile.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if Assigned(FSuccProc) then
  begin
    FSuccProc();
    FSuccProc := nil;
  end;
end;

procedure TfrmMultipleRunsFile.btnPathButtonOnClick(Sender: TObject);
var
  lFileName: string;
begin
  case TButton(Sender).Tag of
    1: // Folder Path
      begin
        with SelectDirectoryDialog do
        begin
          Title := 'Select the folder where the multiple runs files are stored...';
          FileName := '';

          if Execute then
          begin
            lFileName := ExtractFileDir(SelectDirectoryDialog.FileName);

            pathedit.Text := IncludeTrailingPathDelimiter(lFileName + '\');
            addpath.Enabled := true;
          end;
        end;
      end;
    2: // Select Options file
      begin
        SetCurrentDir(pathedit.Text);

        with OpenDialog do
        begin
          InitialDir := '.';
          Title := 'Select the program options file...';
          FileName := '';
          Filter := 'OPTIONS files (*.options)|*.options|TXT files (*.txt)|*.txt|All files (*.*)|*.*';
          FilterIndex := 1;

          if Execute then
            OptionEdit.Text := ExtractFileName(FileName);
        end;
      end;
    3: // Select Item Control File
      begin
        SetCurrentDir(pathedit.Text);

        with OpenDialog do
        begin
          InitialDir := '.';
          FileName := '';
          Filter := 'TXT files (*.txt)|*.txt|DAT files (*.dat)|*.dat|All files (*.*)|*.*';
          FilterIndex := 1;

          if Execute then
            controledit.Text := ExtractFileName(FileName);
        end;
      end;
    4: // Select FDATA File
      begin
        SetCurrentDir(pathedit.Text);

        with OpenDialog do
        begin
          InitialDir := '.';
          Title := 'Select the data file...';
          FileName := '';
          Filter := 'TXT files (*.txt|*.txt|DAT files (*.dat)|*.dat|All files (*.*)|*.*';
          FilterIndex := 1;

          if Execute then
            dataedit.Text := ExtractFileName(FileName);
        end;
      end;
  end;
end;

procedure TfrmMultipleRunsFile.btnOptionsOnClick(Sender: TObject);
begin
  case TButton(Sender).Tag of
    1: // btnAddPath
      begin
        LineNumber := Succ(LineNumber);
        currentMRF.lines.Add('PATH' + #9 + pathedit.Text);

        AddOption.Enabled := true;
        UseDefault.Enabled := true;
      end;
    2: // btnAddOptions
      begin
        if FileExists(pathedit.Text + OptionEdit.Text) then
        begin
          LineNumber := Succ(LineNumber);
          currentMRF.lines.Add('OPTS' + #9 + OptionEdit.Text);

          AddControl.Enabled := true;
          SkipControl.Enabled := true;
        end
        else
          TDialogService.ShowMessage('The specified options file does not exist' + #13 +
            'Please check the Windows PATH and the File Name and try again');
      end;
    3: // btnUseDefaults
      begin
        LineNumber := Succ(LineNumber);
        currentMRF.lines.Add('OPTS' + #9 + 'DEFAULTS');

        AddControl.Enabled := true;
        SkipControl.Enabled := true;
      end;
    4: // AddControl
      begin
        if FileExists(pathedit.Text + controledit.Text) then
        begin
          LineNumber := Succ(LineNumber);
          currentMRF.lines.Add('CTRL' + #9 + controledit.Text);

          AddData.Enabled := true;
        end
        else
          TDialogService.ShowMessage('The specified control file does not exist' + #13 +
            'Please check the Windows PATH and the File Name and try again');
      end;
    5: // SkipControl
      begin
        LineNumber := Succ(LineNumber);
        currentMRF.lines.Add('CTRL' + #9 + 'ITEMAN 3');

        AddData.Enabled := true;
      end;
    6: // btnAddData
      begin
        if FileExists(pathedit.Text + dataedit.Text) then
        begin
          LineNumber := Succ(LineNumber);
          currentMRF.lines.Add('DATA' + #9 + dataedit.Text);

          btnRunMRF.Enabled := true;
        end
        else
          TDialogService.ShowMessage('The specified data file does not exist' + #13 +
            'Please check the Windows PATH and the File Name and try again');
      end;
  end;
end;

procedure TfrmMultipleRunsFile.btnBackClick(Sender: TObject);
begin
  FormResult := mrCancel;
  Close;
end;

procedure TfrmMultipleRunsFile.btnRunClick(Sender: TObject);
begin
  with SaveDialog do
  begin
    Title := 'Please enter a file name for the MRF file ...';
    InitialDir := '.';
    FileName := '';
    Filter := 'TXT files (*.txt|*.txt|DAT files (*.dat)|*.dat|All files (*.*)|*.*';
    FilterIndex := 1;

    if Execute then
    begin
      MRFFile := changefileext(FileName, '.txt');

      if MRFFile <> '' then
      begin
        currentMRF.lines.savetofile(MRFFile);
        TDialogService.ShowMessage('The MRF file ' + MRFFile + ' was saved successfully.');
      end;
    end;
  end;
end;

procedure TfrmMultipleRunsFile.btnRunMRFClick(Sender: TObject);
begin
  RunMRF1(pathedit.Text + changefileext(dataedit.Text, '') + 'MRF.txt', true);
end;

procedure TfrmMultipleRunsFile.RunMRF1(aMRFFileName: string = ''; aSaveMRF: Boolean = False);
var
  i: Integer;
  lName1: string;
  lList: TStrings;
begin
  PathOk := False;
  CTRLOk := False;
  OPTSOk := False;
  DataOk := False;
  MRFOk := False;
  MRFRun := False;

  if aSaveMRF then
    currentMRF.lines.savetofile(aMRFFileName);

  lList := TStringList.Create;
  try
    lList.NameValueSeparator := #9;
    lList.LoadFromFile(aMRFFileName);

    for i := 0 to pred(lList.Count) do
    begin
      if lList.Names[i] = 'PATH' then
      begin
        Path := lList.ValueFromIndex[i];
        PathOk := true;
        CTRLOk := False;
        DataOk := False;
      end
      else if lList.Names[i] = 'OPTS' then
      begin
        OPTS := lList.ValueFromIndex[i];
        OPTSOk := true;
      end
      else if lList.Names[i] = 'CTRL' then
      begin
        CTRL := lList.ValueFromIndex[i];
        CTRLOk := true;
        DataOk := False;
      end
      else if lList.Names[i] = 'DATA' then
      begin
        FDATA := lList.ValueFromIndex[i];
        CTRLOk := true;
        DataOk := true;
      end;

      MRFOk := PathOk and OPTSOk and CTRLOk and DataOk;

      if MRFOk then
        MRFRun := true;
      // if any of the mrf runs were successful, this stays TRUE unlike mrfok

      if MRFOk = true then
      begin
        if OPTS <> 'DEFAULTS' then
          frmMain.ReadOptions(Path + OPTS);
        if OPTS = 'DEFAULTS' then
          frmMain.ReadOptions(frmMain.apppath + '\' + 'Defaults.options');
        if CTRL = 'ITEMAN 3' then
          frmMain.cbDataMatrixFile.IsChecked := true
        else
          frmMain.cbDataMatrixFile.IsChecked := False;

        currentMRF.lines.Clear;
        currentMRF.lines.Add('Analyzing the dataset ' + FDATA);
        frmMain.RunIteman(Path + FDATA, Path + CTRL, Path + changefileext(FDATA, '') + '.docx',
          Path + 'MRF_Summary.txt');
      end;
    end;
  finally
    lList.free;
  end;

  lName1 := Path + 'MRF_Summary.txt';
  frmMain.mrfcall(lName1, MRFRun);
end;

end.
