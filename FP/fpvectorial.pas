unit fpvectorial;

interface

uses
  System.Classes, System.SysUtils, System.Math, System.TypInfo, System.contnrs,
  System.types, System.UITypes,
  // FCL-Image
  fpcanvas, fpimage,

  FMX.Graphics;

const
  { Default extensions }
  { Multi-purpose document formats }
  STR_PDF_EXTENSION = '.pdf';
  STR_POSTSCRIPT_EXTENSION = '.ps';
  STR_SVG_EXTENSION = '.svg';
  STR_SVGZ_EXTENSION = '.svgz';
  STR_CORELDRAW_EXTENSION = '.cdr';
  STR_WINMETAFILE_EXTENSION = '.wmf';
  STR_AUTOCAD_EXCHANGE_EXTENSION = '.dxf';
  STR_ENCAPSULATEDPOSTSCRIPT_EXTENSION = '.eps';
  STR_LAS_EXTENSION = '.las';
  STR_LAZ_EXTENSION = '.laz';
  STR_RAW_EXTENSION = '.raw';
  STR_MATHML_EXTENSION = '.mathml';
  STR_ODG_EXTENSION = '.odg';
  STR_ODT_EXTENSION = '.odt';
  STR_DOCX_EXTENSION = '.docx';
  STR_HTML_EXTENSION = '.html';

  STR_FPVECTORIAL_TEXT_HEIGHT_SAMPLE = 'Ćą';

  NUM_MAX_LISTSTYLES = 8;  // OpenDocument Limit is 10, MS Word Limit is 9

  // Convenience constant to convert text size points to mm
  FPV_TEXT_POINT_TO_MM = 0.35278;

  TWO_PI = 2.0 * pi;

type
  TvVectorialFormat = (
    vfUnknown,
    { Multi-purpose document formats }
    vfPDF, vfSVG, vfSVGZ, vfCorelDrawCDR, vfWindowsMetafileWMF, vfODG,
    { CAD formats }
    vfDXF,
    { Geospatial formats }
    vfLAS, vfLAZ,
    { Printing formats }
    vfPostScript, vfEncapsulatedPostScript,
    { GCode formats }
    vfGCodeAvisoCNCPrototipoV5, vfGCodeAvisoCNCPrototipoV6,
    { Formula formats }
    vfMathML,
    { Text Document formats }
    vfODT, vfDOCX, vfHTML,
    { Raster Image formats }
    vfRAW
    );

  TvPageFormat = (vpA4, vpA3, vpA2, vpA1, vpA0);

  TvProgressEvent = procedure (APercentage: Byte) of object;

  {@@ This routine is called to add an item of caption AStr to an item
    AParent, which is a pointer to another item as returned by a previous call
    of this same proc. If AParent = nil then it should add the item to the
    top of the tree. In all cases this routine should return a pointer to the
    newly created item.
  }
  TvDebugAddItemProc = function (AStr: string; AParent: Pointer): Pointer of object;

//  TvCustomVectorialWriter = class;
//  TvCustomVectorialReader = class;
//  TvPage = class;
//  TvVectorialPage = class;
//  TvTextPageSequence = class;
//  TvEntity = class;
//  TPath = class;
  TvVectorialDocument = class;
//  TvEmbeddedVectorialDoc = class;
//  TvRenderer = class;

  { Coordinates }

  T2DPoint = record
    X, Y: Double;
  end;
  P2DPoint = ^T2DPoint;

  T3DPoint = record
    X, Y, Z: Double;
  end;
  P3DPoint = ^T3DPoint;

  T2DPointsArray = array of T2DPoint;
  T3DPointsArray = array of T3DPoint;
  TPointsArray = array of TPoint;

  { Pen, Brush and Font }

  TvPen = record
    Color: TFPColor;
    Style: TFPPenStyle;
    Width: Integer;
    Pattern: array of LongWord;
  end;
  PvPen = ^TvPen;

  TvBrushKind = (bkSimpleBrush, bkHorizontalGradient, bkVerticalGradient,
    bkOtherLinearGradient, bkRadialGradient);
  TvCoordinateUnit = (vcuDocumentUnit, vcuPercentage);

  TvGradientFlag = (gfRelStartX, gfRelStartY, gfRelEndX, gfRelEndY, gfRelToUserSpace);
  TvGradientFlags = set of TvGradientFlag;

  TvGradientColor = record
    Color: TFPColor;
    Position: Double;   // 0 ... 1
  end;

  TvGradientColors = array of TvGradientColor;

  TvBrush = record
    Color: TFPColor;
    Style: TFPBrushStyle;
    Kind: TvBrushKind;
    Image: TFPCustomImage;
    // Gradient filling support
    Gradient_start: T2DPoint; // Start/end point of gradient, in pixels by default,
    Gradient_end: T2DPoint;   // but if gfRel* in flags relative to entity boundary or user space
    Gradient_flags: TvGradientFlags;
    Gradient_cx, Gradient_cy, Gradient_r, Gradient_fx, Gradient_fy: Double;
    Gradient_cx_Unit, Gradient_cy_Unit, Gradient_r_Unit, Gradient_fx_Unit, Gradient_fy_Unit: TvCoordinateUnit;
    Gradient_colors: TvGradientColors;
  end;
  PvBrush = ^TvBrush;

  TvFont = record
    Color: TFPColor;
    Size: integer;
    Name: utf8string;
    {@@
      Font orientation is measured in degrees and uses the
      same direction as the LCL TFont.orientation, which is counter-clockwise.
      Zero is the normal, horizontal, orientation, directed to the right.
    }
    Orientation: Double;
    Bold: boolean;
    Italic: boolean;
    Underline: boolean;
    StrikeThrough: boolean;
  end;
  PvFont = ^TvFont;

  TvSetStyleElement = (
    // Pen, Brush and Font
    spbfPenColor, spbfPenStyle, spbfPenWidth,
    spbfBrushColor, spbfBrushStyle, spbfBrushGradient, spbfBrushKind,
    spbfFontColor, spbfFontSize, spbfFontName, spbfFontBold, spbfFontItalic,
    spbfFontUnderline, spbfFontStrikeThrough, spbfAlignment,
    // TextAnchor
    spbfTextAnchor,
    // Page style
    sseMarginTop, sseMarginBottom, sseMarginLeft, sseMarginRight
    );

  TvSetStyleElements = set of TvSetStyleElement;
  // for backwards compatibility, obsolete
  TvSetPenBrushAndFontElement = TvSetStyleElement;
  TvSetPenBrushAndFontElements = TvSetStyleElements;

  TvStyleKind = (
    // Paragraph kinds
    vskTextBody, vskHeading,
    // Text-span kind
    vskTextSpan);

  TvStyleAlignment = (vsaLeft, vsaRight, vsaJustifed, vsaCenter);

  TvTextAnchor = (vtaStart, vtaMiddle, vtaEnd);

  { TvStyle }

  TvStyle = class
  protected
    FExtraDebugStr: string;
  public
    Name: string;
    Parent: TvStyle; // Can be nil
    Kind: TvStyleKind;
    Alignment: TvStyleAlignment;
    HeadingLevel: Integer;
    //
    Pen: TvPen;
    Brush: TvBrush;
    Font: TvFont;
    TextAnchor: TvTextAnchor;
    // Page style
    MarginTop, MarginBottom, MarginLeft, MarginRight: Double; // in mm
    SuppressSpacingBetweenSameParagraphs : Boolean;
    //
    SetElements: TvSetStyleElements;
    //
//    Constructor Create;

{    function GetKind: TvStyleKind; // takes care of parenting
    procedure Clear(); virtual;
    procedure CopyFrom(AFrom: TvStyle);
////    procedure CopyFromEntity(AEntity: TvEntity);
    procedure ApplyOverFromPen(APen: PvPen; ASetElements: TvSetStyleElements);
    procedure ApplyOverFromBrush(ABrush: PvBrush; ASetElements: TvSetStyleElements);
    procedure ApplyOverFromFont(AFont: PvFont; ASetElements: TvSetStyleElements);
    procedure ApplyOver(AFrom: TvStyle); virtual;
////    procedure ApplyIntoEntity(ADest: TvEntity); virtual;
    function CreateStyleCombinedWithParent: TvStyle;
    function GenerateDebugTree(ADestRoutine: TvDebugAddItemProc; APageItem: Pointer): Pointer; virtual;
}  end;

  TvListStyleKind = (vlskBullet, vlskNumeric);

  TvNumberFormat = (vnfDecimal,      // 0, 1, 2, 3...
                    vnfLowerLetter,  // a, b, c, d...
                    vnfLowerRoman,   // i, ii, iii, iv....
                    vnfUpperLetter,  // A, B, C, D...
                    vnfUpperRoman);  // I, II, III, IV....
  { TvListLevelStyle }

  TvListLevelStyle = Class
    Kind : TvListStyleKind;
    Level : Integer;
    Start : Integer; // For numbered lists only

    // Define the "leader", the stuff in front of each list item
    Prefix : String;
    Suffix : String;
    Bullet : String; // Only applies to Kind=vlskBullet
    NumberFormat : TvNumberFormat; // Only applies to Kind=vlskNumeric
    DisplayLevels : Boolean; // Only applies to numbered lists.
                             // If true, style is 1.1.1.1.
                             //     else style is 1.
    LeaderFontName : String; // Not used by odt...

    MarginLeft : Double; // mm
    HangingIndent : Double; //mm
    Alignment : TvStyleAlignment;

//    Constructor Create;
  end;

 { TvListStyle }

  TvListStyle = class
  private
    ListLevelStyles : TList;
  public
    Name : String;

{    constructor Create;
    destructor Destroy; override;

    procedure Clear;
    function AddListLevelStyle : TvListLevelStyle;
    function GetListLevelStyleCount : Integer;
    function GetListLevelStyle(AIndex: Integer): TvListLevelStyle;
}  end;

  { ... }

  { TvVectorialDocument }

  TvVectorialDocument = class
  private
    FOnProgress: TvProgressEvent;
    FPages: TList;
    FStyles: TList;
    FListStyles: TList;
    FCurrentPageIndex: Integer;
////    FRenderer: TvRenderer;
////    function CreateVectorialWriter(AFormat: TvVectorialFormat): TvCustomVectorialWriter;
////    function CreateVectorialReader(AFormat: TvVectorialFormat): TvCustomVectorialReader;
  public
    Width, Height: Double; // in millimeters
    Name: string;
    Encoding: string; // The encoding on which to save the file, if empty UTF-8 will be utilized. This value is filled when reading
    ForcedEncodingOnRead: string; // if empty, no encoding will be forced when reading, but it can be set to a LazUtils compatible value
    // User-Interface information
    ZoomLevel: Double; // 1 = 100%
    { Selection fields }
////    SelectedElement: TvEntity;
    // List of common styles, for conveniently finding them
    StyleTextBody, StyleHeading1, StyleHeading2, StyleHeading3,
      StyleHeading4, StyleHeading5, StyleHeading6: TvStyle;
    StyleTextBodyCentralized, StyleTextBodyBold: TvStyle; // text body modifications
    StyleHeading1Centralized, StyleHeading2Centralized, StyleHeading3Centralized: TvStyle; // heading modifications
    StyleBulletList, StyleNumberList : TvListStyle;
    StyleTextSpanBold, StyleTextSpanItalic, StyleTextSpanUnderline: TvStyle;
    // Reader properties
////    ReaderSettings: TvVectorialReaderSettings;
    { Base methods }

    constructor Create; virtual;
    destructor Destroy; override;
    procedure Assign(ASource: TvVectorialDocument);
    procedure AssignTo(ADest: TvVectorialDocument);
(*    procedure WriteToFile(AFileName: string; AFormat: TvVectorialFormat); overload;
    procedure WriteToFile(AFileName: string); overload;
    procedure WriteToStream(AStream: TStream; AFormat: TvVectorialFormat);
    procedure WriteToStrings(AStrings: TStrings; AFormat: TvVectorialFormat);
    procedure ReadFromFile(AFileName: string; AFormat: TvVectorialFormat); overload;
    procedure ReadFromFile(AFileName: string); overload;
    procedure ReadFromStream(AStream: TStream; AFormat: TvVectorialFormat);
    procedure ReadFromStrings(AStrings: TStrings; AFormat: TvVectorialFormat);
////    procedure ReadFromXML(ADoc: TXMLDocument; AFormat: TvVectorialFormat);
    class function GetFormatFromExtension(AFileName: string; ARaiseException: Boolean = True): TvVectorialFormat;
    function  GetDetailedFileFormat(): string;
    procedure GuessDocumentSize();
    procedure GuessGoodZoomLevel(AScreenSize: Integer = 500);
    { Page methods }
//    function GetPage(AIndex: Integer): TvPage;
//    function GetPageIndex(APage : TvPage): Integer;
//    function GetPageAsVectorial(AIndex: Integer): TvVectorialPage;
//    function GetPageAsText(AIndex: Integer): TvTextPageSequence;
    function GetPageCount: Integer;
//    function GetCurrentPage: TvPage;
//    function GetCurrentPageAsVectorial: TvVectorialPage;
    procedure SetCurrentPage(AIndex: Integer);
    procedure SetDefaultPageFormat(AFormat: TvPageFormat);
//    function AddPage(AUseTopLeftCoords: Boolean = False): TvVectorialPage;
//    function AddTextPageSequence(): TvTextPageSequence;
    { Style methods }
    function AddStyle(): TvStyle;
    function AddListStyle: TvListStyle;
    procedure AddStandardTextDocumentStyles(AFormat: TvVectorialFormat);
    function GetStyleCount: Integer;
    function GetStyle(AIndex: Integer): TvStyle;
    function FindStyleIndex(AStyle: TvStyle): Integer;
    function GetListStyleCount: Integer;
    function GetListStyle(AIndex: Integer): TvListStyle;
    function FindListStyleIndex(AListStyle: TvListStyle): Integer;
    { Data removing methods }
    procedure Clear; virtual;
    { Drawer selection methods }
////    function GetRenderer: TvRenderer;
////    procedure SetRenderer(ARenderer: TvRenderer);
    procedure ClearRenderer();
    { Debug methods }
    procedure GenerateDebugTree(ADestRoutine: TvDebugAddItemProc; APageItem: Pointer = nil);
    { Events }
    property OnProgress: TvProgressEvent read FOnProgress write FOnprogress;      *)
  end;

  { ... }

implementation


constructor TvVectorialDocument.Create;
begin
  inherited Create;

  FPages := TList.Create;
  FCurrentPageIndex := -1;
  FStyles := TList.Create;
  FListStyles := TList.Create;
//  if gDefaultRenderer <> nil then
//    FRenderer := gDefaultRenderer.Create;
end;

destructor TvVectorialDocument.Destroy;
begin
////  Clear();

  FPages.Free;
  FPages := nil;
  FStyles.Free;
  FStyles := nil;
  FListStyles.Free;
  FListStyles := nil;

////  ClearRenderer();

  inherited Destroy;
end;

procedure TvVectorialDocument.Assign(ASource: TvVectorialDocument);
//var
//  i: Integer;
begin
//  Clear;
//
//  for i := 0 to ASource.GetEntitiesCount - 1 do
//    Self.AddEntity(ASource.GetEntity(i));
end;

procedure TvVectorialDocument.AssignTo(ADest: TvVectorialDocument);
begin
  ADest.Assign(Self);
end;

end.
