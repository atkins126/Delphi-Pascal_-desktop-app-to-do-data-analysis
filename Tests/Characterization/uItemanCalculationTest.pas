unit uItemanCalculationTest;

interface

uses
  DUnitX.TestFramework,
  uItemanCalculation;

type

  [TestFixture]
  TItemanCalculationTest = class
  private
    FTestee: TItemanCalculation;
  public
    [Setup]
    procedure Setup;
    [TearDown]
    procedure TearDown;
    [Test]
    procedure Test_RunIteman_SmokeTest;
  end;

implementation

uses
  ufrmMain;

procedure TItemanCalculationTest.Setup;
begin
  FTestee := TItemanCalculation.Create([]);
  frmMain := TfrmMain.Create(nil);
end;

procedure TItemanCalculationTest.TearDown;
begin
  frmMain.Free;
  frmMain := Pointer(1);
  // not nil but some intentionally wrong value to cause errors
  FTestee.Free;
  FTestee := Pointer(1); // see above
end;

procedure TItemanCalculationTest.Test_RunIteman_SmokeTest;
begin
  Assert.IsTrue(FTestee.RunIteman(3));
end;

initialization

TDUnitX.RegisterTestFixture(TItemanCalculationTest);

end.
