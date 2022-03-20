program Iteman;

uses
  System.StartUpCopy,
  FMX.Forms,
  ufrmMain in 'GUI\ufrmMain.pas' {frmMain} ,
  ufrmMultipleRunsFile in 'GUI\ufrmMultipleRunsFile.pas' {frmMultipleRunsFile} ,
  vkbdhelper in 'ORM\vkbdhelper.pas',
  ufrmLicense in 'GUI\ufrmLicense.pas' {frmLicense} ,
  uProtect in 'Classes\uProtect.pas',
  uItemanExport in 'Classes\uItemanExport.pas',
  uItemanCalculation in 'Classes\uItemanCalculation.pas';

{$R *.res}
{$R ItemanResource.RES}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;

end.
