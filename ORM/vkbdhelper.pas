unit vkbdhelper;

{
  Force focused control visible when Android/IOS Virtual Keyboard showed or hiden
  How to use:
  place vkdbhelper into your project uses section. No more code needed.

  Changes by ZuBy
  ======= =======
  2016.01.24
  * clean Uses section
  * top margin value for Virtual Keyboard (global var VKOffset := [integer])
  * now cross-platform (IOS, Android)

  Changes
  =======
  2015.7.12
  * Fix space after hide ime and rotate
  * Fix rotate detection

}
interface

implementation

uses
  System.Classes, System.SysUtils, System.Types, System.Messaging,
  FMX.Types, FMX.Controls, FMX.Layouts, FMX.Forms, FMX.VirtualKeyboard;

type
  TVKStateHandler = class(TComponent)
  protected
    FVKMsgId: Integer;
    FSizeMsgId: Integer;
    FLastControl: TControl;

    FLastMargin: TPointF;
    FLastAlign: TAlignLayout;
    FLastBounds: TRectF;

    FVKVisibleTimer: TTimer;
    procedure DoVKVisibleChanged(const Sender: TObject; const Msg: System.Messaging.TMessage);
    procedure DoSizeChanged(const Sender: TObject; const Msg: System.Messaging.TMessage);
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure DoVKVisibleCheck(ASender: TObject);
    procedure EnableVKCheck(AEnabled: Boolean);
  public
    constructor Create(AOwner: TComponent); overload; override;
    destructor Destroy; override;
  end;

var
  VKHandler: TVKStateHandler;
  FglVKB: TRectF;
  VKOffset: Integer = 20;

function GetVKBounds(var ARect: TRect): Boolean;
var
  ContentRect, TotalRect: TRectF;
begin
  TotalRect := Screen.ActiveForm.ClientRect;
  ContentRect.Bottom := (TotalRect.Bottom - FglVKB.Top);
  Result := TotalRect.Bottom <> ContentRect.Bottom;
  if Result then
  begin
    ARect.Left := trunc(TotalRect.Left);
    ARect.Top := trunc(ContentRect.Bottom);
    ARect.Right := trunc(TotalRect.Right);
    ARect.Bottom := trunc(TotalRect.Bottom);
  end;
end;

function IsVKVisible: Boolean;
var
  R: TRect;
begin
  Result := GetVKBounds(R);
end;

{ TVKStateHandler }
constructor TVKStateHandler.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FVKMsgId := TMessageManager.DefaultManager.SubscribeToMessage(TVKStateChangeMessage, DoVKVisibleChanged);
  FSizeMsgId := TMessageManager.DefaultManager.SubscribeToMessage(TSizeChangedMessage, DoSizeChanged);
  FVKVisibleTimer := TTimer.Create(Self);
  FVKVisibleTimer.Enabled := False;
  FVKVisibleTimer.Interval := 100;
  FVKVisibleTimer.OnTimer := DoVKVisibleCheck;
end;

destructor TVKStateHandler.Destroy;
begin
  TMessageManager.DefaultManager.Unsubscribe(TVKStateChangeMessage, FVKMsgId);
  TMessageManager.DefaultManager.Unsubscribe(TSizeChangedMessage, FSizeMsgId);
  inherited;
end;

procedure TVKStateHandler.DoSizeChanged(const Sender: TObject; const Msg: System.Messaging.TMessage);
var
  ASizeMsg: TSizeChangedMessage absolute Msg;
  R: TRect;
  AScene: IScene;
  AScale: Single;
begin
  if Sender = Screen.ActiveForm then
  begin
    if GetVKBounds(R) then
    begin
      if Assigned(FLastControl) then
      begin
        if FLastControl is TLayout then
        begin
          if FLastControl.TagObject = Self then
          begin
            FLastControl.SetBounds(0, 0, Screen.ActiveForm.Width, Screen.ActiveForm.Height);
          end;
        end
        else //
          TCustomScrollBox(FLastControl).Margins.Bottom := 0;
      end;
      if Supports(Sender, IScene, AScene) then
      begin
        AScale := AScene.GetSceneScale;
        R.Left := trunc(R.Left / AScale);
        R.Top := trunc(R.Top / AScale);
        R.Right := trunc(R.Right / AScale);
        R.Bottom := trunc(R.Bottom / AScale);
        TMessageManager.DefaultManager.SendMessage(Sender, TVKStateChangeMessage.Create(true, R));
      end;
    end
  end;
end;

procedure TVKStateHandler.DoVKVisibleChanged(const Sender: TObject; const Msg: System.Messaging.TMessage);
var
  AVKMsg: TVKStateChangeMessage absolute Msg;
  ACtrl: TControl;
  ACtrlBounds, AVKBounds, ATarget: TRectF;

  procedure MoveCtrls(AOldParent, ANewParent: TFmxObject);
  var
    I: Integer;
    AChild: TFmxObject;
  begin
    I := 0;
    while I < AOldParent.ChildrenCount do
    begin
      AChild := AOldParent.Children[I];
      if AChild <> ANewParent then
      begin
        if AChild.Parent = AOldParent then
        begin
          AChild.Parent := ANewParent;
          Continue;
        end;
      end;
      Inc(I);
    end;
  end;
  procedure AdjustByLayout(R: TRectF; ARoot: TFmxObject);
  var
    ALayout: TLayout;
  begin
    if (ARoot.ChildrenCount = 1) and (ARoot.Children[0] is TLayout) then
    begin
      ALayout := ARoot.Children[0] as TLayout;
      if ALayout.Align <> TAlignLayout.None then
      begin
        FLastAlign := ALayout.Align;
      end;
      FLastBounds := ALayout.BoundsRect;
      FLastMargin.Y := ALayout.Position.Y;
      FLastMargin.X := ALayout.Position.X;
      ALayout.Align := TAlignLayout.None;
    end
    else
    begin
      ALayout := TLayout.Create(ARoot);
      ALayout.Parent := ARoot;
      ALayout.TagObject := Self;
      MoveCtrls(ARoot, ALayout);
      FLastMargin.Y := 0;
      FLastMargin.X := 0;
      FLastAlign := TAlignLayout.Client;
      FLastBounds := TRectF.Create(0, 0, 0, 0);
    end;
    if ARoot is TForm then
    begin
      if (TForm(ARoot).Width <> ALayout.Width) or (TForm(ARoot).Height <> ALayout.Height) then
        ALayout.SetBounds(0, 0, TForm(ARoot).Width, TForm(ARoot).Height);
    end;
    ALayout.Position.Y := R.Bottom - ACtrlBounds.Bottom;
    if FLastControl <> ALayout then
    begin
      if Assigned(FLastControl) then
        FLastControl.RemoveFreeNotification(Self);
      FLastControl := ALayout;
      FLastControl.FreeNotification(Self);
      EnableVKCheck(true);
    end;
  end;

  procedure ScrollInToRect(R: TRectF);
  var
    AParent, ALastParent: TFmxObject;
    AParentBounds: TRectF;
    AScrollBox: TCustomScrollBox;
    AOffset: Single;
  begin
    AParent := ACtrl.Parent;
    AScrollBox := nil;
    ALastParent := AParent;
    while Assigned(AParent) do
    begin
      if AParent is TCustomScrollBox then
      begin
        AScrollBox := AParent as TCustomScrollBox;
        AParentBounds := AScrollBox.AbsoluteRect;
        if AParentBounds.Contains(R) then
        begin
          AOffset := ACtrlBounds.Top - R.Top + R.Height+30;
          if (AParentBounds.Bottom > AVKBounds.Top) or (AParentBounds.Bottom < AParentBounds.Height) then
          begin
            if (FLastControl <> AScrollBox) then
            begin
              if Assigned(FLastControl) then
                FLastControl.RemoveFreeNotification(Self);
              FLastMargin.Y := AScrollBox.Margins.Bottom;
              FLastMargin.X := AScrollBox.Margins.Left;
              FLastControl := AScrollBox;
              FLastControl.FreeNotification(Self);
              EnableVKCheck(true);
            end;
            AScrollBox.Margins.Bottom := AParentBounds.Bottom - AVKBounds.Top+30;
          end;
          AScrollBox.ViewportPosition := TPointF.Create(AScrollBox.ViewportPosition.X, AScrollBox.ViewportPosition.Y + AOffset);
          Break;
        end;
      end;
      ALastParent := AParent;
      AParent := AParent.Parent;
    end;
    if not Assigned(AScrollBox) then
      AdjustByLayout(R, ALastParent);
  end;

begin
  if AVKMsg.KeyboardVisible then
  begin
    if Screen.FocusControl <> nil then
    begin
      ACtrl := Screen.FocusControl.GetObject as TControl;
      ACtrlBounds := ACtrl.AbsoluteRect;
      AVKBounds := TRectF.Create(AVKMsg.KeyboardBounds);
      FglVKB := AVKBounds;
      AVKBounds.Top := AVKBounds.Top - VKOffset;
      if (ACtrlBounds.Bottom > AVKBounds.Top) or (ACtrlBounds.Top < 0) then
      begin
        ATarget := ACtrlBounds;
        ATarget.Top := AVKBounds.Top - ACtrlBounds.Height;
        ATarget.Bottom := ATarget.Top + ACtrlBounds.Height;
        ScrollInToRect(ATarget);
      end
    end
  end
  else
  begin
    FglVKB := TRectF.Empty;
    if Assigned(FLastControl) then
    begin
      if FLastControl is TCustomScrollBox then
      begin
        FLastControl.Margins.Bottom := FLastMargin.Y;
        FLastControl.Margins.Left := FLastMargin.X;
      end
      else
      begin
        if FLastAlign = TAlignLayout.None then
        begin
          FLastControl.Position.Y := FLastMargin.Y;
          FLastControl.Position.X := FLastMargin.X;
        end
        else
        begin
          FLastControl.BoundsRect := FLastBounds;
          FLastControl.Align := FLastAlign;
        end;
      end;
      FLastControl := nil;
      EnableVKCheck(False);
    end;
  end;
end;

procedure TVKStateHandler.DoVKVisibleCheck(ASender: TObject);
{ procedure FMXAndroidFix;
  var
  AService: IFMXVirtualKeyboardService;
  AVK: TVirtualKeyboardAndroid;
  AListener: TVKListener;
  AContext: TRttiContext;
  AType: TRttiType;
  AField: TRttiField;
  begin
  if TPlatformServices.Current.SupportsPlatformService(IFMXVirtualKeyboardService, AService) then
  begin
  AVK := AService as TVirtualKeyboardAndroid;
  AContext := TRttiContext.Create;
  AType := AContext.GetType(AVK.ClassType);
  if Assigned(AType) then
  begin
  AField := AType.GetField('FVKListener');
  if Assigned(AField) then
  begin
  AListener := AField.GetValue(AVK).AsObject as TVKListener;
  AListener.onVirtualKeyboardHidden;
  end;
  end;
  Screen.ActiveForm.Focused := nil;
  end;
  end; }

begin
  if not IsVKVisible then
  begin
    EnableVKCheck(False);
    if Assigned(Screen.FocusControl) then
      TMessageManager.DefaultManager.SendMessage(Screen.FocusControl.GetObject,
        TVKStateChangeMessage.Create(False, TRect.Create(0, 0, 0, 0)));
    // FMXAndroidFix;
  end;
end;

procedure TVKStateHandler.EnableVKCheck(AEnabled: Boolean);
begin
  FVKVisibleTimer.Enabled := AEnabled;
end;

procedure TVKStateHandler.Notification(AComponent: TComponent; Operation: TOperation);
begin
  if Operation = opRemove then
  begin
    if FLastControl = AComponent then
      FLastControl := nil;
  end;
  inherited;
end;

initialization

FglVKB := TRectF.Empty;
VKHandler := TVKStateHandler.Create(nil);

finalization

FreeAndNil(VKHandler);

end.
