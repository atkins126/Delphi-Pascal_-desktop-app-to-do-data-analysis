unit ufrmLicense;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,

  vkbdhelper, uProtect,

  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs, FMX.Layouts,
  FMX.StdCtrls, FMX.ScrollBox, FMX.Memo, FMX.Edit, FMX.Objects,
  FMX.Controls.Presentation;

type
  TfrmLicense = class(TForm)
    LayoutContent: TLayout;
    pBG: TPanel;
    gplMain: TGridPanelLayout;
    Layout12: TLayout;
    gplTopInfoBlock: TGridPanelLayout;
    Layout13: TLayout;
    gplDataContent: TGridPanelLayout;
    Layout1: TLayout;
    gplTab1Row1: TGridPanelLayout;
    Layout26: TLayout;
    Text2: TText;
    Layout27: TLayout;
    Layout3: TLayout;
    GridPanelLayout1: TGridPanelLayout;
    Layout4: TLayout;
    Text1: TText;
    Layout5: TLayout;
    eMachineKey: TEdit;
    Layout8: TLayout;
    GridPanelLayout2: TGridPanelLayout;
    Layout9: TLayout;
    Text3: TText;
    Layout10: TLayout;
    eLicenseKey: TEdit;
    Layout15: TLayout;
    GridPanelLayout3: TGridPanelLayout;
    Layout16: TLayout;
    Text4: TText;
    Layout18: TLayout;
    gplBottom: TGridPanelLayout;
    Layout42: TLayout;
    btnRequestLicenseKey: TButton;
    Layout43: TLayout;
    btnConfirmLicense: TButton;
    tbMain: TToolBar;
    btnBack: TButton;
    tLicenseStatus: TText;
    Layout17: TLayout;
    Layout2: TLayout;
    rbLicType365: TRadioButton;
    rbLicType180: TRadioButton;
    GoodLogo: TImage;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    FSuccProc: TProc;
    FFormResult: TModalResult;

    FProtect: TProtect;
  public
    procedure RunForm(const aSuccProc: TProc);

    property FormResult: TModalResult read FFormResult write FFormResult;
    property Protect: TProtect read FProtect write FProtect;
  end;

  function ShowLicenseForm(): TfrmLicense;

var
  frmLicense: TfrmLicense;

implementation

uses
  ufrmMain;

{$R *.fmx}

function ShowLicenseForm(): TfrmLicense;
begin
  frmLicense:= TfrmLicense.Create(Application);

  with frmLicense do
  begin
    tbMain.StylesData['Caption.Text']:= 'Iteman License Activation';

    tLicenseStatus.Text:= 'Demo License';

    FProtect:= TProtect.GetProtect;
    if Assigned(FProtect) then
    begin
      tLicenseStatus.Text:= FProtect.GetLicenseStatusDes(False);
      eMachineKey.Text:= FProtect.Key;
      eLicenseKey.Text:= FProtect.License;

      rbLicType180.IsChecked:= FProtect.LicenseType = lct30;
      rbLicType365.IsChecked:= FProtect.LicenseType = lct365;
    end;
  end;
  Result:= frmLicense;
end;

procedure TfrmLicense.RunForm(const aSuccProc: TProc);
begin
  FSuccProc:= aSuccProc;
  FormResult:= mrCancel;
  Show;
end;

procedure TfrmLicense.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if Assigned(FSuccProc) then
  begin
    FSuccProc();
    FSuccProc:= nil;
  end;
end;

end.
