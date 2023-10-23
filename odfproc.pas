{**
  English version
  This unit contains classes and methods which gives you possibility to work
  with OpenDocument Format. This means that you can create, edit and save 
  .odt files (.ods and other not realized yet).

  @author Vladislav Baghenov (http://www.webdelphi.ru)
  @author of patch: leo
  @author of patch: v-t-l
  @author of patch: Rustam Asmandiarov (http://predatorglscene.blogspot.com/)

  @Lisence LGPL

------------------------------------------------------------------------------------
 
  Русская версия
  Этот модуль содержит классы и методы позволяющие работать с форматом 
  OpenDocument. Это значит что вы можете создавать, редактировать и сохранять
  файлы .odt, также реализована начальная поддержка .ods, поддержка
  других форматов пока не реализована).

  @автор Владислав Баженов (http://www.webdelphi.ru)
  @автор дополнений: leo
  @автор дополнений: v-t-l
  @автор дополнений: Рустам Асмандияров (http://predatorglscene.blogspot.com/)

  @Лицензия LGPL

}  

unit ODFProc;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Graphics, XMLRead, XMLWrite, DOM, FileUtil, LResources,
  Forms, Dialogs, dateutils, LCLProc, process, zipper, Math, fgl, LCLIntf;

const
  DefaultOdsFileName = 'doc.ods';
  DefaultOdtFileName = 'doc.odt';

  Win32_Cmd ='C:/Windows/System32/cmd.exe';
  DocEntrys: array [1..3] of string = ('content.xml','meta.xml','styles.xml');

  xmlns: array [1..20,1..2] of string =(('office','urn:oasis:names:tc:opendocument:xmlns:office:1.0'),
                                        ('style' ,'urn:oasis:names:tc:opendocument:xmlns:style:1.0'),
                                        ('text'  ,'urn:oasis:names:tc:opendocument:xmlns:text:1.0'),
                                        ('table' ,'urn:oasis:names:tc:opendocument:xmlns:table:1.0'),
                                        ('draw'  ,'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0'),
                                        ('fo'    ,'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0'),
                                        ('meta'  ,'urn:oasis:names:tc:opendocument:xmlns:meta:1.0'),
                                        ('number','urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0'),
                                        ('svg'   ,'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0'),
                                        ('chart' ,'urn:oasis:names:tc:opendocument:xmlns:chart:1.0'),
                                        ('dr3d'  ,'urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0'),
                                        ('form'  ,'urn:oasis:names:tc:opendocument:xmlns:form:1.0'),
                                        ('script','urn:oasis:names:tc:opendocument:xmlns:script:1.0'),
                                        ('of'    ,'urn:oasis:names:tc:opendocument:xmlns:of:1.2'),
                                        ('field' ,'urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0'),
                                        ('formx' ,'urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0'),
                                        ('dc'    ,'http://purl.org/dc/elements/1.1/'),
                                        ('xlink' ,'http://www.w3.org/1999/xlink'),
                                        ('math'  ,'http://www.w3.org/1998/Math/MathML'),
                                        ('xforms','http://www.w3.org/2002/xforms'));

{ Массив для именования/разыменования автостилей }
  symarr: array [1..52] of char = ('A','B','C','D','E','F','G','H','I','J',
  'K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','a','b','c',
  'd','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v',
  'w','x','y','z');


//type // простой тип пары данных индекс-значение
//  TCouple = record
//    Index: integer;
//    Value: string;
//  end;
//  TCouples = array of TCouple; // массив пар данных индекс-значение

type
  TFontStyle = (ftBold, ftItalic, ftUnderline);
  TFontStyles = set of TFontStyle;
  TTextPosition = (tpCenter, tpLeft, tpRight, tpJustify);
  TVertAlign = (vaTop, vaMiddle, vaBottom, vaAutomatic);
  TSizeCounter = (tsPercent, tsCm); //размерность ширины элемента таблицы (проценты или см.)
  TFileType = (ftStyles, ftContent, ftManifest, ftMeta,ftSettings);
  TStrArr = array of string;

  TPositiveInteger = object
    private
      FValue: Integer;
    end;
  operator := (val: Integer) r: TPositiveInteger;
  operator := (val: TPositiveInteger) r: Integer;
  operator = (val1, val2: TPositiveInteger) r: boolean;

type
  TNonNegativeInteger = object
    private
      FValue: Integer;
    end;
  operator := (val: Integer) r: TNonNegativeInteger;
  operator := (val: TNonNegativeInteger) r: Integer;
  operator = (val1, val2: TNonNegativeInteger) r: boolean;

  //TNonNegativeLength = object

{ Расширенный булев формат }
const
  ExtBoolArr: array [0..2] of string = ('', 'true', 'false');
type
  // расширенный булев формат, с маркером nil для обозначения неустановленной
  // величины
  TExtBool = (ebNil, ebTrue, ebFalse);


{ Поддерживаемые форматы файлов документов }
const
  SupportedExtensionsArr: array [0..6] of string = ('.pdf','.doc','.odt',
                                                    '.xls','.ods', '.ppt',
                                                    '.odp');
type
  TSupportedExtensions = (sePDF, seDOC, seODT, seXLS, seODS, sePPT, seODP);


const
  ValueTypesArr: array [0..7] of string = ('boolean', 'currency', 'date',
                              'float', 'percentage', 'string', 'time', 'void');
type
  // тип описывает типы данных
  TValueTypes = (vtBoolean, vtCurrency, vtDate, vtFloat, vtPercentage,
                                                     vtString, vtTime, vtVoid);

{ Единицы измерения }
const
  MeasureArr: array [0..8] of string = ('cm', 'mm', 'in', 'pt', 'pc', 'px',
                                                               'em', '%', '*');
type
  // тип едениц измерения
  TMeasure = (mCm, mMm, mIn, mPt, mPc, mPx, mEm, mPercent, mRelative);

{ Направление письма }
const
  // виды письма слава-направо, сверху-вниз и тд.
  WritingModeArr: array [0..7] of string = ('lr-tb','rl-tb','tb-rl','tb-lr','lr','rl','tb','page');
type
  TWritingMode = (wm_lrtb, wm_rltb, wm_tbrl, wm_tblr, wm_lr, wm_rl, wm_tb, wm_page);

{ Разрыв страницы }
const
  BreakArr: array [0..3] of string = ('', 'auto', 'column', 'page');
type
  TBreak = (bNil, bAuto, bColumn, bPage);

{ Выравнивание таблицы }
const
  TableAlignArr: array [0..3] of string = ('center', 'left', 'right', 'margins');
type
  TTableAlign = (taCenter, taLeft, taRight, taMargins);

{ Прилипание к следующему абзацу }
const
  KeepWithNextArr: array [0..1] of string = ('auto', 'always');
type
  TKeepWithNext = (knAuto, knAlways);

{ Модель границы таблицы }
const
  BorderModelArr: array [0..1] of string = ('collapsing', 'separating');
type
  TBorderModel = (bmCollapsing, bmSeparating);

{ Видимость колонки/строки }
const
  TableVisibilityArr: array [0..2] of string = ('collapse', 'filter', 'visible');
type
  TTableVisibility = (tvCollapse, tvFilter, tvVisible);

{ Переполнение текста }
const
  WrapOptionArr: array [0..1] of string = ('no-wrap', 'wrap');
type
  TWrapOption = (woNoWrap, woWrap);

{ Варианты защиты ячейки }
const
  CellProtectArr: array [0..3] of string = ('none', 'hidden-and-protected',
                                            'protected', 'formula-hidden');
type
  TCellProtectValues = set of (cpNone, cpHiddenAndProtected, cpProtected,
                                                             cpFormulaHidden);

type
  { TCellProtect }
  TCellProtect = class
  private
    FValue: TCellProtectValues;
    procedure SetValue(AValue: TCellProtectValues);
  public
    property Value: TCellProtectValues read FValue write SetValue;
  end;

{ Направление символов }
const
  DirectionArr: array [0..1] of string = ('ltr', 'ttb');
type
  TDirection = (dLeftToRight, dTopToBottom);

{ TGlyphOrientationVertical }
const
  GlyphOrientationVerticalArr: array [0..4] of string = ('auto', '0', '0deg',
                                                              '0rad', '0grad');
type
  TGlyphOrientationVertical = (govAuto, gov0, gov0Deg, gov0Rad, gov0Grad);

{ Выравнивание при вращении }
const
  RotationAlignArr: array [0..3] of string = ('none', 'bottom', 'top', 'center');
type
  TRotationAlign = (raNone, raBottom, raTop, raCenter);

{ Источник значения выравнивания ячейки }
const
  TextAlignSourceArr: array [0..1] of string = ('fix', 'value-type');
type
  TTextAlignSource = (tasFix, tasValueType);

{ Вертикальное выравнивание в ячейке }
const
  CellVerticalAlignArr: array [0..3] of string = ('top', 'middle', 'bottom', 'automatic');
type
  TCellVerticalAlign = (cvaTop, cvaMiddle, cvaBottom, cvaAutomatic);

const DefaultTablePosition = taMargins;
      DefaultTableShadow   = '#808080 0.18cm 0.18cm';
      DefaultTableMargin   = 1;
      DefaultTableName     = 'Таблица';
      DefaultSizeCounter   = tsPercent;
      DefaultVertAlign     = vaMiddle;
      DefaultTableWidth    = 100;
      DefaultTableStyle    = '_Table';
      DefaultColStyle      = '_Column';
      DefaultCellStyle     = '_Cell';
      DefaultCellBorder    = '0.002cm solid #000000';
      DefaultOptimalWidth  = true;
      DefaultTextStyle     = 'Standard';


type
  TOdtTableStyle = record
    Name       : string;        //название стиля
    SizeCounter: TSizeCounter;  //размерность % или см.
    Width      : currency;      //ширина таблицы
    Align      : TTableAlign;//положение таблицы на странице
    Margin     : integer;       //отступы
end;

type
  TOdtTableColRowStyle = record
    Name: string;
    SizeCounter: TSizeCounter;   //размерность % или см
    ColWidth      : currency;     //ширина/высота столюца/строки
    RowHeight     : currency;
    UseOptimalColWidth: boolean; //использовать оптимальную ширину столбца
    UseOptimalRowWidth: boolean; //использовать оптимальную высоту строки
end;

type
  TOdtTableCellStyle = record
    Name: string;
    VerticalAlign : TVertAlign;
    border: string;
    border_top: string;
    border_bottom: string;
    border_left: string;
    border_right: string;
end;

type
  TLength = object      // размерный тип
    private
      FValue: double;     // значение
      FMeasure: TMeasure; // единица измерения
    public
      property Value: double read FValue write FValue;
      property Measure: TMeasure read FMeasure write FMeasure;
    end;

type // тип "неотрицательной длины"
  TNonNegativeLength = object(TLength)
    private
      procedure SetValue(AValue: double);
    public
      property Value: double read FValue write SetValue;
    end;

type // тип "положительной длины"
  TPositiveLength = object(TLength)
    private
      procedure SetValue(AValue: double);
    public
      property Value: double read FValue write SetValue;
    end;

type
  TPageNumber = record
    Value: TPositiveInteger;
    Auto: boolean;
    BreakNumber: TDOMnode;
  end;

  {
      9.1.4 <table:table-cell>
      The <table:table-cell> element represents a table cell. It is contained in a table row
      element. A table cell can contain paragraphs and other text content as well as sub tables. Table
      cells may span multiple columns and rows. Table cells may be empty.
      The <table:table-cell> element is usable within the following element: <table:table-
      row> 9.1.3.

      The <table:table-cell> element has the following attributes:
      office:boolean-value 19.367,
      office:currency 19.369,
      office:date-value 19.370,
      office:string-value 19.379,
      office:time-value 19.382,
      office:value 19.384,
      office:value-type 19.385,
      table:content-validation-name 19.601,
      table:formula 19.642,
      table:number-columns-repeated 19.675.3,
      table:number-columns-spanned 19.676,
      table:number-matrix-columns-spanned 19.679,
      table:number-matrix-rows-spanned 19.680,
      table:number-rows-spanned 19.678,
      table:protect 19.695,
      table:protected 19.696.5,
      table:style-name 19.726.13,
      xhtml:about 19.905,
      xhtml:content 19.906,
      xhtml:datatype 19.907,
      xhtml:property 19.908,
      xml:id 19.914.

      The <table:table-cell> element has the following child elements:
      <dr3d:scene> 10.5.2,
      <draw:a> 10.4.12,
      <draw:caption> 10.3.11,
      <draw:circle> 10.3.8,
      <draw:connector> 10.3.10,
      <draw:control> 10.3.13,
      <draw:custom-shape> 10.6.1,
      <draw:ellipse> 10.3.9,
      <draw:frame> 10.4.2,
      <draw:g> 10.3.15,
      <draw:line> 10.3.3,
      <draw:measure> 10.3.12,
      <draw:page-thumbnail> 10.3.14,
      <draw:path> 10.3.7,
      <draw:polygon> 10.3.5,
      <draw:polyline> 10.3.4,
      <draw:rect> 10.3.2,
      <draw:regular-polygon> 10.3.6,
      <office:annotation> 14.1,
      <table:cell-range-source> 9.3.1,
      <table:detective> 9.3.2,
      <table:table> 9.1.2,
      <text:alphabetical-index> 8.8,
      <text:bibliography> 8.9,
      <text:change> 5.5.7.4,
      <text:change-end> 5.5.7.3,
      <text:change-start> 5.5.7.2,
      <text:h> 5.1.2,
      <text:illustration-index> 8.4,
      <text:list> 5.3.1,
      <text:numbered-paragraph> 5.3.6,
      <text:object-index> 8.6,
      <text:p> 5.1.3,
      <text:section> 5.4,
      <text:soft-page-break> 5.6,
      <text:table-index> 8.5,
      <text:table-of-content> 8.3,
      <text:user-index> 8.7.
  }

type
  { TOdtTableCell }
  TOdtTableCell = class
  private
    FBooleanValue: boolean;
    FCurrency: string;
    FDateValue: TDateTime;
    FProtectedCell: boolean;
    FStringValue: string;
    //FTimeValue                                                  не реализовано
    FValue: double;
    FValueType: TValueTypes;  // определяет тип значения ячейки
    //FContentValidationName: string;                             не реализовано
    //FFormula: string;                                           не реализовано
    FNumberColumnsRepeated: TPositiveInteger; // во скольких следующих колонках ячейка повторится
    FNumberColumnsSpanned: TPositiveInteger;  // сколько следующих ячеек ячейка объединит
    FNumberMatrixColumnsSpanned: TPositiveInteger; // число объединенных колонок
    FNumberMatrixRowsSpanned: TPositiveInteger; // число объединенных строк
    FNumberRowsSpanned: TPositiveInteger; // сколько следующих строк ячейка объединит
    //FProtect                                                    не реализовано
    FProtected: boolean; // false - можно редактировть true - нельзя
    FStyleName: string; // стиль ячейки
    procedure SetBooleanValue(AValue: boolean);
    procedure SetCurrency(AValue: string);
    procedure SetDateValue(AValue: TDateTime);
    procedure SetProtectedCell(AValue: boolean);
    procedure SetStringValue(AValue: string);
    procedure SetStyleName(AValue: string);
    procedure SetValue(AValue: double);
    procedure SetValueType(AValue: TValueTypes);
    //FAbout                                                      не реализовано
    //FContent                                                    не реализовано
    //FDatatype                                                   не реализовано
    //FProperty                                                   не реализовано
    //FXMLId: string;                                             не реализовано
  public
    property BooleanValue: boolean read FBooleanValue write SetBooleanValue;
    property Currency: string read FCurrency write SetCurrency;
    property DateValue: TDateTime read FDateValue write SetDateValue;
    property StringValue: string read FStringValue write SetStringValue;
    property Value: double read FValue write SetValue;
    property ValueType: TValueTypes read FValueType write SetValueType;  // определяет тип значения ячейки
    property NumberColumnsRepeated: TPositiveInteger read FNumberColumnsRepeated write FNumberColumnsRepeated; // во скольких следующих колонках ячейка повторится
    property NumberColumnsSpanned: TPositiveInteger read FNumberColumnsSpanned write FNumberColumnsSpanned;  // сколько следующих ячеек ячейка объединит
    property NumberMatrixColumnsSpanned: TPositiveInteger read FNumberMatrixColumnsSpanned write FNumberMatrixColumnsSpanned; // число объединенных колонок
    property NumberMatrixRowsSpanned: TPositiveInteger read FNumberMatrixRowsSpanned write FNumberMatrixRowsSpanned; // число объединенных строк
    property NumberRowsSpanned: TPositiveInteger read FNumberRowsSpanned write FNumberRowsSpanned; // сколько следующих строк ячейка объединит
    property ProtectedCell: boolean read FProtectedCell write SetProtectedCell; // false - можно редактировть true - нельзя
    property StyleName: string read FStyleName write SetStyleName; // стиль ячейки
  end;

type
  TOdtTableCells = Specialize TFPGList<TOdtTableCell>; // массив ячеек

  {
  The <style:table-cell-properties> element has the following attributes:
  fo:background-color 20.175,
  fo:border 20.176.2,
  fo:border-bottom 20.176.3,
  fo:border-left 20.176.4,
  fo:border-right 20.176.5,
  fo:border-top 20.176.6,
  fo:padding 20.210,
  fo:padding-bottom 20.211,
  fo:padding-left 20.212,
  fo:padding-right 20.213,
  fo:padding-top 20.214,
  fo:wrap-option 20.223,
  style:border-line-width 20.241,
  style:border-line-width-bottom 20.242,
  style:border-line-width-left 20.243,
  style:border-line-width-right 20.244,
  style:border-line-width-top 20.245,
  style:cell-protect 20.246,
  style:decimal-places 20.250,
  style:diagonal-bl-tr 20.251,
  style:diagonal-bl-tr-widths 20.252,
  style:diagonal-tl-br 20.253,
  style:diagonal-tl-br-widths 20.254,
  style:direction 20.255,
  style:glyph-orientation-vertical 20.289,
  style:print-content 20.323.3,
  style:repeat-content 20.334,
  style:rotation-align 20.338,
  style:rotation-angle 20.339,
  style:shadow 20.349,
  style:shrink-to-fit 20.350,
  style:text-align-source 20.354,
  style:vertical-align 20.386.2
  style:writing-mode 20.394.6.

  The <style:table-cell-properties> element has the following child element:
  <style:background-image> 17.3.
  }

type
  { TOdtTableCellProperties }
  TOdtTableCellProperties = class
  private
    FBackgroundColor: TColor;  // цвет ячейки
    FBorder: string; // граница
    FBorderBottom: string; // нижняя граница
    FBorderLeft: string; // левая граница
    FBorderRight: string; // правая граница
    FBorderTop: string; // верхняя граница
    FPadding: TNonNegativeLength; // внутренний отступ
    FPaddingBottom: TNonNegativeLength; // внутренний отступ снизу
    FPaddingLeft: TNonNegativeLength; // внутренний отступ слева
    FPaddingRight: TNonNegativeLength; // внутренний отступ справа
    FPaddingTop: TNonNegativeLength; // внутренний отступ сверху
    FWrapOption: TWrapOption; // скрывать/показывать текст выходящий за границу ячейки
    FBorderLineWidth: TPositiveLength; // ширина границы               ! различный способ отображения при различных моделях границ
    FBorderLineWidthBottom: TPositiveLength; // ширина границы снизу   ! различный способ отображения при различных моделях границ
    FBorderLineWidthLeft: TPositiveLength; // ширина границы слева     ! различный способ отображения при различных моделях границ
    FBorderLineWidthRight: TPositiveLength; // ширина границы справа   ! различный способ отображения при различных моделях границ
    FBorderLineWidthTop: TPositiveLength; // ширина границы сверху     ! различный способ отображения при различных моделях границ
    FCellProtect: TCellProtect; // описывает как защищена ячейка
    FDecimalPlaces: TNonNegativeInteger;
    FDiagonalBLTR: string;
    FDiagonalBLTRWidths: TPositiveLength;
    FDiagonalTLBR: string;
    FDiagonalTLBRWidths: TPositiveLength;
    FDirection: TDirection;         // направление символов
    FGlyphOrientationVertical: TGlyphOrientationVertical;
    FPrintContent: boolean; // печатается ли контент
    FRepeatContent: boolean; // повторяется ли контент
    FRotationAlign: TRotationAlign;  // выравнивание при вращении
    //FRotationAngle      // угол вращения, требуется раелизация типа
    //FShadow             // тень
    FShrinkToFit: boolean;
    FTextAlignSource: TTextAlignSource;
    FVerticalAlign: TCellVerticalAlign;
    FWritingMode: TWritingMode; // атр, режим письма (слева-направо, сверху-вниз и т.д.)
  public
//
  end;

type
  TOdtTableCellsProperties = Specialize TFPGList<TOdtTableCellProperties>; // массив свойств ячеек


type
  { TOdtTableLineProperties }
  TOdtTableLineProperties = class
  protected
  private
    //FBreakAfter: TBreak;  // атр, разрыв страницы после колонки/строки
    //FBreakBefore: TBreak; // атр, разрыв страницы до колонки/строки
    //FUseOptimalSize: TExtBool; // атр, оптимальная ширина/высота колонки/строки
    function Equal(Obj: TOdtTableLineProperties): boolean;
  public
    //property BreakAfter: TBreak read FBreakAfter write FBreakAfter;
    //property BreakBefore: TBreak read FBreakBefore write FBreakBefore;
    //property UseOptimalSize: TExtBool read FUseOptimalSize write FUseOptimalSize;
  end;

type
  { TOdtTableColumnProperties }
  TOdtTableColumnProperties = class(TOdtTableLineProperties)
  private
    FSymbol: string;      // часть значения атрибута - символьное обозначение колонки A,B,C...
    FWidth: TLength;       // атр, абс определяет фикс ширину колонки
    //FRelWidth: TLength;    // атр, относительная ширина колонки
  public
    function Equal(Obj: TOdtTableColumnProperties): boolean;
    property Symbol: string read FSymbol write FSymbol;
    property Width: TLength read FWidth write FWidth;
    //property RelWidth: TLength read FRelWidth write FRelWidth;
  end;

type
  TOdtTableColsProperties = Specialize TFPGList<TOdtTableColumnProperties>; // массив свойств колонок


  {
    The <style:table-row-properties> element has the following attributes:
    fo:background-color
    fo:break-after
    fo:break-before
    fo:keep-together
    style:min-row-height
    style:row-height
    style:use-optimal-row-height
    The <style:table-row-properties> element has the following child element:
    <style:background-image> 17.3.
  }

type
  { TOdtTableRowProperties }
  TOdtTableRowProperties = class(TOdtTableLineProperties)
  private
    FBackgroundColor: TColor;  // цвет строки
    FKeepTogether: TKeepWithNext; // атр, не отрывать строку от соседних?
    FMinHeight: TLength;        // атр, минимальная высота строки
    FHeight: TLength;           // атр, высота строки
    //FBackgroundImage  // не реализовано
  public
    function Equal(Obj: TOdtTableRowProperties): boolean;
    property BackgroundColor: TColor read FBackgroundColor write FBackgroundColor default 0;
    property KeepTogether: TKeepWithNext read FKeepTogether write FKeepTogether;
    property MinHeight: TLength read FMinHeight write FMinHeight;
    property Height: TLength read FHeight write FHeight;
  end;
type
  TOdtTableRowsProperties = Specialize TFPGList<TOdtTableRowProperties>; // массив свойств строк


type
  { TOdtTableLine }
  TOdtTableLine = class
  private
    FDefaultCellStyleName: string;
    //FNumberRepeated: integer; // свойство разворачивается - вставляется указанное количество объектов
    FStyleName: string;
    FVisibility: TTableVisibility;
    //FXMLId: string;       не реализовано
  public
    function Equal(Obj: TOdtTableLine): boolean;
    property DefaultCellStyleName: string read FDefaultCellStyleName write FDefaultCellStyleName;
    //property NumberRepeated: integer read FNumberRepeated write FNumberRepeated;
    property StyleName: string read FStyleName write FStyleName;
    property Visibility: TTableVisibility read FVisibility write FVisibility;
    //property XMLId: string read FXMLId write FXMLId;
  end;

type

  { TOdtTableColumn }

  TOdtTableColumn = class(TOdtTableLine);                  // тип колонки
  TOdtTableColumns = Specialize TFPGList<TOdtTableColumn>; // тип массива колонок
  TOdtTableRow = class(TOdtTableLine);                     // тип строки
  TOdtTableRows = Specialize TFPGList<TOdtTableRow>;       // тип массива строк

type
  { TOdtTableProperties }
  TOdtTableProperties = class
  private
    FBackgroundColor: TColor; // атр, цвет фона, 'transparent' или '#rrggbb'  --- работает
    //FBackgroundImage:   // дочерняя нода, картинка фона
    FBreakAfter: TBreak;  // атр, разрыв страницы после таблицы      --- работает
    FBreakBefore: TBreak; // атр, разрыв страницы до таблицы         --- работает
    FKeepWithNext: TKeepWithNext;   // атр, не отрывать от следующего абзаца   --- работает
    //FMargin: TLength;       // атр, отступ, неотрицательное абс или отн (%) значение --- не работает в ООо3.2, вместо него используются конкретные значения отступа - слева, справа...
    FMarginBottom: TNonNegativeLength; // атр, отступ снизу, неотрицательное абс или отн (%) значение
    FMarginLeft: TLength;   // атр, отступ слева, абс или отн (%) значение, мб отрицательным --- работает
    FMarginRight: TLength;  // атр, отступ справа, абс или отн (%) значение, мб отрицательным
    FMarginTop: TNonNegativeLength;    // атр, отступ сверху, неотрицательное абс или отн (%) значение
    FMayBreakBetweenRows: boolean;  // атр, разрыв таблицы для переноса на след страницу --- работает
    FPageNumber: TPageNumber;    // атр, номер страницы после разрыва таблицы --- работает
    FRelWidth: TLength;    // атр, ширина таблицы относительно ширины поля (%) в которм находится таблица --- работает
    FShadow: string;         // атр, тень, --- пока не реализован класс для ввода параметров --- работает
    FWidth: TLength;       // атр, абс определяет фикс ширину таблицы (cm)  --- работает
    //FWritingMode: TWritingMode; // атр, режим письма (слева-направо, сверху-вниз и т.д.) --- не работает в OOo3.2 и AbiWord, вместо этого используется такой же атрибут для каждой ячейки
    FAlign: TTableAlign;  // атр, горизонтальное выравнивание таблицы  --- работает
    FBorderModel: TBorderModel; // атр, модель границ таблицы          --- работает
                               // (в ООо - галочка 'объединить стили смежных линий')
    //FDisplay: boolean;       // атр, видимость таблицы, есть в спецификации, но  --- не работает в ООо3.2
                            // атрибут Display похоже не поддерживается редакторами
    procedure SetWidth(w:TLength);
    procedure SetRelWidth(w:TLength);
  public
    property BackgroundColor: TColor read FBackgroundColor write FBackgroundColor default 0;
    //property BackgroundImage:
    property BreakAfter: TBreak read FBreakAfter write FBreakAfter;
    property BreakBefore: TBreak read FBreakBefore write FBreakBefore;
    property KeepWithNext: TKeepWithNext read FKeepWithNext write FKeepWithNext;
    //property Margin: TLength read FMargin write SetMargin;
    property MarginBottom: TNonNegativeLength read FMarginBottom write FMarginBottom;
    property MarginLeft: TLength read FMarginLeft write FMarginLeft;
    property MarginRight: TLength read FMarginRight write FMarginRight;
    property MarginTop: TNonNegativeLength read FMarginTop write FMarginTop;
    property MayBreakBetweenRows: boolean read FMayBreakBetweenRows write FMayBreakBetweenRows;
    property PageNumber: TPageNumber read FPageNumber write FPageNumber;
    property RelWidth: TLength read FRelWidth write SetRelWidth;
    property Shadow: string read FShadow write FShadow;
    property Width: TLength read FWidth write SetWidth;
    //property WritingMode: TWritingMode read FWritingMode write FWritingMode;
    property Align: TTableAlign read FAlign write FAlign;
    property BorderModel: TBorderModel read FBorderModel write FBorderModel;
    //property Display: boolean read FDisplay write FDisplay;
  end;

type
  { TPage }
  TOdt = class;
  TOdtTable = class;
type
  TPage = class(TObject)
  private
    FParent : TOdt;
    FPreviousPage : TDOMNode;
    FBreakNumber : TDOMNode;
    FNextPage : TDOMNode;
  public
    procedure FindAndReplace(Search,Replace:string;ReplaceOnce: boolean=false);//поиск и замена текста
    //поиск встречающегося слова в тексте.
    //IncludeChildren-ищет искомое слово в сразу в дочернем узле а не только в конкретном)
    function FindText(Search : UTF8string; IncludeChildren: boolean=true) : UTF8string;
    function FindNode(Search : UTF8string; IncludeChildren: boolean=true) : TDOMNode;
    // procedure FindAndRemove(Search:string);//поиск и удаление
    function AppendText(aNode: TDOMNode; StyleName, Text: string):boolean;
    function AppendSpace(aNode: TDOMNode; CountSpaces: integer):boolean;

    function InsertTable(Cols, Rows:integer;TableName:string):TOdtTable;
    function InsertTable(InsertinBefore : TDOMnode; Cols, Rows : integer; TableName : string) : TOdtTable;

    procedure DeleteTable(TableName:string); overload;
    procedure DeleteTable(aTable: TOdtTable); overload;
    function  TableExists(TableName: string) : boolean; //В этом листе такая таблица сущществует?
    function  GetTable(TableName: string):TOdtTable;// опубликована для получения таблицы по названию
    function  GetListofTables: TStrings;// опубликована для получения списка

    property PreviousPage: TDOMNode read FPreviousPage;
    property BreakNumber: TDOMnode read FBreakNumber;
    property NextPage: TDOMNode read FNextPage;
  end;

  { TOdfTable }
  { Родитель классов TOdsSheet и TOdtTable }
type
  TOdfTable = class
  private
    function GetRowNode(ARow: Integer): TDOMNode;
  protected
    FName     : string;       //имя таблицы атрибут table:name
    FDocument : TXMLDocument; //документ в котором расположена таблица
    FRoot     : TDOMNode;     //корневая нода table:table
  public
    constructor Create(XMLDoc: TXMLDocument; TableName: string);
    destructor Destroy; override;
    //поиск и замена текста Search на Replace в строке ARow таблицы шаблона
    function FindAndReplace(ARow: integer; Search, Replace: string): boolean;
    //клонирование строки таблицы ARow
    procedure MultiplyRow(ARow, Count: Integer);
  end;

  { TOdf }
  { Родитель классов TOdt TOds }
type
  TOdf = class
  private // свойства и методы доступны только внутри класса
    function ExportScriptToOO: boolean;
    function GetMetaAuthor: string;
    function GetMetaGenerator: string;
    function GetOfficeVersion: string;
    procedure SetMetaAuthor(const Author: string);
    procedure SetMetaGenerator(const AGenerator: string);
  protected // свойства и методы доступны только для дочерних классов
    FStyles   : TXMLDocument; //styles.xml - стили текста, форматирования таблиц и т.д.
    FContent  : TXMLDocument; //content.xml - содержимое документа
    FManifest : TXMLDocument; //META-INF/manifest.xml - содержимое всего архива
    FMeta     : TXMLDocument; //meta.xml - мета-данные (автор, дата генерации документа, генератор и т.д.)
    FSettings : TXMLDocument; //Еще один файл в котором лежат общие настройки документа OpenOffice
    FTempDir  : string;       // Папка с временными файлами
    FRoot : TDOMnode;         //Корень дерева данных
    FName : string;

    procedure GenerateContent;
    procedure GenerateManifest;
    procedure GenerateStyles;
    procedure GenerateMeta; // получаем основную таблицу
    procedure InsertXMLNS(var RootNode: TDOMElement);

  public
    constructor Create;
    destructor Destroy; override;

    property Styles      : TXMLDocument read FStyles write FStyles;
    property Content     : TXMLDocument read FContent write FContent;
    property Manifest    : TXMLDocument read FManifest write FManifest;
    property Meta        : TXMLDocument read FMeta write FMeta;
    property Settings    : TXMLDocument read FSettings write FSettings;
    property TempDir     : string read FTempDir write FTempDir;

  public // свойства и методы доступные для изменения, в том числе у объектов-потомков

    // закрывает документ
    function CloseDocument: boolean;

    // проверяет загружен ли документ
    function DocumentLoaded: boolean;

    // загружает компонент документа (Содержимое/Манифест/Стили/Мета)
    function LoadPartOfDocument(FileName: string; Doc: TFileType): boolean;

    // загружает шаблон из файла
    function LoadFromFile(FileName: string): boolean;

    // генерация документа
    procedure GenerateDocument(DocumentName: string;
                         const DocumentPath: string='default');

    // просмотр получившегося документа
    procedure ShowDocument(DocumentName: string; Editor: string='default');

    // вывод документа на принтер
    function PrintDocument(FileName: string): boolean;

    // заменяет текст всех дочерних нод FRoot с Search на Replace
    function FindAndReplace(Search, Replace: string): boolean;

    // удаляёт ноду содержащую текст Search
    function FindAndRemove(Search: string): boolean;

    // возвращает таблицу по названию
    function GetTable(TableName: string): TOdfTable;

    // конвертирует документ в нужный формат
    // конвертирование требует установленного Libre\OpenOffice
    // тестирование проводилось на LibreOffice 3 / ubuntu 11 и LibreOffice 4 / Debian 7
    // список поддерживаемых форматов содержится в массиве SupportedExtensionsArr,
    // протестированные направления конвертации:
    // .odt => .doc; .odt => .pdf; .ods => .xls
    // вероятно также будут работать конвертации:
    // .doc => .odt; .doc => .pdf; .xls => .ods; .xls => .pdf
    // но они не тестировались и требуют более подробного рассмотрения перед использованием
    // под Windows не работает   (с)Leo
    function ConvertTo(InFileName: string; to_ext: TSupportedExtensions;
      var OutFileName: string): boolean; overload;
    function ConvertTo(InFileName: string; to_ext_str: string;
      var OutFileName: string): boolean; overload;

    // название программы создавшей документ
    property Generator   : string read GetMetaGenerator write SetMetaGenerator;

    // автор документа
    property Author      : string read GetMetaAuthor write SetMetaAuthor;
  end;

type
  { TOdtTable }
  TOdtTable = class(TOdfTable)
  private
    FDefTextStyle: string; //стиль по умолчанию для текста таблицы
    {новые свойства класса}
    FProperties : TOdtTableProperties; // свойства таблицы
    FColsProperties: TOdtTableColsProperties; // свойства колонок
    FColumns: TOdtTableColumns; // колонки таблицы
    FRows: TOdtTableRows; // строки таблицы
    FCells: TOdtTableCells; // ячейки таблицы
    procedure GetColsProperties; // получить свойства колонок таблицы
    function GetSymbol(s: string): string;
    procedure GetTableProperties(var P: TOdtTableProperties); // получить свойства таблицы
    procedure GetColumns; // получить колонки таблицы
    procedure GetRows; // получить строки таблицы
    function  GetColCount: integer; //получение количества столбцов таблицы
    function  GetRowCount: integer; //получение количества строк таблицы
    function InTable(s: string): boolean; // отвечает принадлежит стиль к таблице или нет
    function GetHeaderRowsNode(ACol: Integer): TDOMNode;
    function  GetHeaderRows(ACol: Integer): string; //получение содержимого ячейки
    procedure SetHeaderRows(ACol: Integer; const AValue: string);//запись данных в ячейку

    procedure SetColCount(AColCount:integer);//установка нового количества столбѲ
    procedure SetRowCount(ARowCount:integer);//установка нового количества строк
    function  GetCells(ACol, ARow: Integer): string; //получение содержимого ячейки
    procedure SetCells(ACol, ARow: Integer; const AValue: string);//запись данных в ячейку
    function GetCellNode(ACol, ARow: Integer):TDOMNode;//поиск узла ячейки
    function CheckFontName(AFontName: string):boolean; //проверка наличия шрифта в документе
    procedure InsertNewFont(AFontName:string);//добавление нового шрифта
    procedure SetName(aName:string);
  public
    constructor Create(XMLDoc:TXMLDocument; TableName:string);
    destructor Destroy; override;
    procedure SetTextStyle(StyleName,FontName:string; FontSize:integer; FontStyles: TFontStyles; TextPosition: TTextPosition);
    procedure ApplyTextStyle(ACol, ARow: Integer; AStyle: string);//применение стиля текста к выбраной ячейке
    procedure AppendColumn(TextStyle:string);//вставка столбца в конец таблицы
    procedure AppendRow(TextStyle:string);//вставка строки в конец таблицы
    procedure InsertRow();//вставка строки со стилями от предедущих строк
    procedure RemoveEmptyRow(Prefix:string);//удаление пустых строк из таблицы
    procedure RemoveRow;Overload;//удаление последней строки
    procedure RemoveRow(Count:Integer);Overload;//удаление Nго количества последних строк
    procedure RemoveColumn;//удаление последнего столбца
    procedure SetTableProperties; // установить свойства таблицы, выполнить после внесённых изменений для записи в хмл
    procedure SetColsProperties; // установить свойства колонок таблицы, выполнить после внесённых изменений для записи в хмл
    procedure SetColumns; // записать изменённые колонки таблицы в хмл
    property Document: TXMLDocument read FDocument;
    property Name: string read FName write SetName;
    property RootNode:TDOMNode read FRoot;
    property ColCount : integer read GetColCount write SetColCount;
    property RowCount : integer read GetRowCount write SetRowCount;
    property DefTextStyle:string read FDefTextStyle write FDefTextStyle;
    property Properties: TOdtTableProperties read FProperties write FProperties; // GetTableProperties write SetTableProperties;
    property ColsProperties: TOdtTableColsProperties read FColsProperties
                      write FColsProperties;
    property Columns: TOdtTableColumns read FColumns write FColumns;
    property Rows: TOdtTableRows read FRows write FRows;
    property Cells[ACol, ARow: Integer]: string read GetCells write SetCells;
    {Заголовок таблицы}
    property HeaderRows[ACol: Integer]: string read GetHeaderRows write SetHeaderRows;
    function HeaderRowsExists(): Boolean;  //Заголовок таблицы сущществует?
  end;

type
  { TOdt }
  //TOdt = class(TObject)
  TOdt = class(TOdf)
  private
    function StylesFindFntFace(FontName: string):TDOMNode;//ищет описание шрифта в styles.xml
    function ContentFindFntFace(FontName: string):TDOMNode;//ищет описание шрифта в content.xml
    procedure GenerateManifest;//создает файл manifest.xml
    procedure GenerateContent;//создает content.xml с основными узлами
    procedure SetDefaultTableStyles(TableName: string);//установка дефолтных стилей для таблицы, столбца и ячейки
    function CheckDefaultTableStyles(TableName: string): boolean;//проверка дефолтных стилей для таблицы
    function GetTablesCount:integer;//получение количества таблиц
    function SetTableStyle(TableStyle:TOdtTableStyle): boolean;
    function SetTableColStyle(TableColRowStyle:TOdtTableColRowStyle):boolean;
    function SetTableCellStyle(TableCellStyle:TOdtTableCellStyle): boolean;

    function GetCountPages():Integer;
    function GetPage(aIndex: integer):TPage;

    procedure RemoveStyleParagraphParametr(Style:String;Parametr:string); //Удаляет какой либо параметр параграфа в стиле
    function GetBeginPages() : TDOMnode;//получаем начало всех страниц в массиве, в самом начале мб быть какой нить мусор
    function GetEndPages() : TDOMnode; //Получаем конец документа
    //Получаем начало таблицы или многоструктурированного дерева.
    //необходимо если разрыв страницы находится внутри таблицы или чего то сложного.
    function GetFixParentNode(OrignNode: TDOMnode) : TDOMnode;
    function ExportScripttoOO(): boolean;
  public
    FCashPages: array of TPageNumber;
    constructor Create;
    destructor Destroy; override;

    function LoadFromFile(FileName: string): boolean;
    function SavetoFile(FileName: string): boolean;//сохранение документа
    function SaveState(): boolean;//Сохранение всех правок в каталоге документа
    property Name: string read FName;
    //   Просмотр получившегося xml
    function ShowPartofDocument(Doc: TFileType = ftContent; editor: string = 'default'): boolean;

    // сборка документа в один файл
    procedure GenerateDocument(DocumentName: string = DefaultOdtFileName;
                                    const DocumentPath: string = 'default');

    // просмотр получившегося документа
    procedure ShowDocument(DocumentName: string=DefaultOdtFileName;
                                    Editor: string='default');

    // Вывод документа на принтер, минуя просмотр в редакторе
    function PrintDocument(DocumentName: string = DefaultOdtFileName): boolean;

    function AddFont(const FontName: string; var Doc: TXMLDocument):boolean;//добавление нового шрифта в styles.xml
    //Стиль текста: Bold,Italic,UnderStyle
    function AddParagraphFontStyle(StyleName: string;
                                   FontStyle: TFontStyles;
                                   FontStyleParent: boolean = false): boolean;overload;
    function AddParagraphFontStyle(FontStyle: TFontStyles;
                                   FontStyleParent: boolean = false): string;overload;
    function GetParagraphFontStyle(FontStyle: TFontStyles): string;
    //Выравнивание текста с лева, по центру,...
    function AddParagraphTextPosition(StyleName: string; TextPosition: TTextPosition):boolean;
    //размер шрифта,использовать ТОЛЬКО для куска текста, парент которого содержит полноценный стиль
    function AddParagraphFontSize(StyleName: string; FontSize:integer): boolean;overload;
    function AddParagraphFontSize(FontSize:integer): string;overload;

    function AddParagraphStyle(StyleName,FontName: string; FontSize:integer; FontStyle: TFontStyles; TextPosition: TTextPosition):boolean;
    function isStyleExists( text: string ): boolean;
    function SearchFreeStyleName(StyleName : string; DontSearchinTable: Boolean=true) : string;

    function AppendText(StyleName, Text: string):boolean;

    // Обработка страниц
    property CountPages:Integer read GetCountPages; //Число страниц
    property Page[aIndex: Integer]: TPage read GetPage;

    {* Для ускорения работы кэшируем страницы,
       рекомендуется обновлять при добавлении новых страниц, или
       после удаления тегов разрыва страниц
       если не кэшировать TPage может работать неправильно
    *}
    procedure CashPages();
    procedure AddPage(aIndex : integer);  //не работает (нужно стиль создавать)
    //Удаление страницы. WithBreakPageonTable=true если разрыв страницы находится в таблице то
    //таблица будет удалена.
    procedure RemovePage(aIndex : integer; WithBreakPageonTable : boolean = true); overload;   //работает
    //Удаление группы страниц. WithBreakPageonTable=true с таблицами в которых разрыв страницы
    procedure RemovePage(aFromIndex, aToIndex : integer; WithBreakPageonTable : boolean = true); overload;   //работает
    //Перемещение страницы.  WithBreakPageonTable=true с таблицами в которых разрыв страницы
    procedure MovePage(Source, Target : integer; WithBreakPageonTable : boolean = true);overload;  //работает
    //Перемещение групп страниц. Например с 1 по 2ю в 3ю.
    procedure MovePage(aFromSource,aToSource, Target : integer; WithBreakPageonTable : boolean = true); overload; //работает
    //Поиск страницы по слову или группе слов разделенных пробелом
    function FindPage(Search : string ): Integer; overload;  //работает(вроде с таблицами тоже)
    //Копирование страницы.  Работает, но могут быть баги
    procedure Copy(Source, Target: Integer; WithBreakPageonTable : boolean = true); overload;
    //Копирование нескольких страниц. Например с 1 по 2ю в 4ю. работает, но могут быть баги
    procedure Copy(aFromSource,aToSource, Target: Integer; WithBreakPageonTable : boolean = true); overload;
    //копируем страницу из другого документа
    {Рустам: Примечание
     1)При копировании документа метод проверяет наличие одноименных стилей
     если таковые имеются создает новый с другим именем(Оптимальным будет создать менеджер стилей)
     2)Не работает перемещение картинок в папках хотя стили перемещяются.
     3)не реализовано перемещение шрифтов
     4)метод крайне тяжелый и многократно использует рекурсии с вложенными циклами,
       для оптимизации нужно будет использовать инлайны КЧ деревья и прочее.
     Заключение: На простом документе работает но как поведет на сложном ничего не могу сказать.
    }
    procedure Copy(SourceDocument: TOdt; Source, Target: Integer; WithBreakPageonTable : boolean = true); overload;

    //обработка таблиц
   function InsertTable(Cols, Rows:integer) : TOdtTable;
    function InsertTable(Cols, Rows:integer; TableName : string) : TOdtTable;
    function InsertTable(InsertinBefore : TDOMnode; Cols, Rows : integer) : TOdtTable;
    function InsertTable(InsertinBefore : TDOMnode; Cols, Rows : integer; TableName : string) : TOdtTable;
    procedure DeleteTable(TableName: string); overload;
    procedure DeleteTable(aTable: TOdtTable); overload;
    function  TableExists(TableName: string) : boolean; //Сущществует ли таблица
    // опубликована для получение таблицы по названию
    function GetTable(TableName: string): TOdtTable;
    function  GetNewTableName() : String;// опубликована для получения таблицы по названию
    function  GetListofTables: TStrings;// для получения списка  таблиц
    property TablesCount : integer read GetTablesCount;

    // поиск текста во всех вложенных нодах
    function FindTextInChildNodes(Node: TDOMNode; Search: UnicodeString): TDOMNode;

  end;

type
  { TOdsSheet }
  TOdsSheet = class(TOdfTable);

type
  { TOds }
  TOds = class(TOdf)
  private
    procedure GenerateContent;
    procedure GenerateManifest;
  public
    constructor Create;
    destructor Destroy; override;

    function LoadFromFile(FileName: string): boolean;
    // генерация документа .ods
    procedure GenerateDocument(DocumentName: string=DefaultOdsFileName;
      const DocumentPath: string='default');
    // просмотр получившегося документа .ods
    procedure ShowDocument(DocumentName: string = DefaultOdsFileName;
                                     Editor: string = 'default');
    // просмотр вывод на принтер получившегося документа .ods
    function PrintDocument(DocumentName: string=DefaultOdsFileName): boolean;
    // возвращает лист по названию
    function GetSheet(SheetName: string) : TOdsSheet;
    function ConvertTo(InFileName: string; ext: string; var OutFileName: string): boolean; overload;
    function ConvertTo(InFileName: string; ext: TSupportedExtensions;
      var OutFileName: string): boolean; overload;
  end;

  function Assembledtext(node: TDOMNode): string;



var ODT: TOdt;

implementation
uses LazUTF8, LazFileUtils, LazUtilities;


//--------------------- TPositiveInteger
operator:=(val: Integer)r: TPositiveInteger;
begin
  if val <= 0 then
    raise EInOutError.Create('Value mast be > 0.')
  else
    r.FValue := val;
end;

operator:=(val: TPositiveInteger)r: Integer;
begin
  r := val.FValue;
end;

operator=(val1, val2: TPositiveInteger) r: boolean;
begin
  if val1.FValue=val2.FValue then r:=true
  else r:=false
end;

//--------------------- TNonNegativeInteger
operator:=(val: Integer)r: TNonNegativeInteger;
begin
  if val < 0 then
    raise EInOutError.Create('Value mast be >= 0.')
  else
    r.FValue := val;
end;

operator:=(val: TNonNegativeInteger)r: Integer;
begin
  r := val.FValue;
end;

operator=(val1, val2: TNonNegativeInteger) r: boolean;
begin
  if val1.FValue=val2.FValue then r:=true
  else r:=false
end;

// преобразование строки в TPageNumber
function StrToPageNumber(AValue: string): TPageNumber;
begin
  Result.Auto := false;
  Result.Value := 1;
  if AValue='auto' then Result.Auto := true
  else Result.Value := StrToInt(AValue);
end;

// преобразование TPageNumber в строку
function PageNumberToStr(AValue: TPageNumber): string;
begin
  if AValue.Auto then Result := 'auto'
  else Result := IntToStr(AValue.Value);
end;

// проверка изменений в полях типа TPageNumber
function Changed(V,OV:TPageNumber):boolean; overload;
begin
  Result := false;
  if (V.Value<>OV.Value) or (V.Auto<>OV.Auto) then Result := true;
end;


function ColorToHtmlHex(Color: TColor):string;
var l: integer;
    sc: string;
begin
  sc:=ColorToString(Color);
  l:=length(sc);
  Result:='#'+copy(sc,l-1,2)+copy(sc,l-3,2)+copy(sc,l-5,2);
end;

function HtmlHexToColor(Color: string):TColor ;
var l: integer;
    cc: string;
begin
  l:=length(Color);
  cc:='$00'+copy(Color,l-1,2)+copy(Color,l-3,2)+copy(Color,l-5,2);
  Result:=StringToColor(cc);
end;

function createUniqueString:string;
var
   MyGuid:TGuid;
begin
     if CreateGUID(MyGuid)=0 then
        begin
             Result := GUIDToString(MyGuid);
             Result := copy(Result,2,Length(Result)-2);
        end
     else
         Result :='UniqueString'+IntToStr(trunc(now));
end;

// перевод символьного значения в цифровое
function SymToNum(s:string):string;
var r,l:integer;
    sum:int64=0;
const c = 52;

  function Idx(t:string):int64;
  var x:integer;
  begin
    for x:=1 to 52 do
      if symarr[x]=t then exit(int64(x));
  end;

begin
  l:=length(s);
  for r:= 1 to length(s) do begin
    sum:=sum+Idx(s[r])*c**(l-r);
  end;
  Result:=IntToStr(sum);
end;

// перевод цифрового значения в символьное
function NumToSym(n:integer):string;
var d,e:integer;
    ds,es:string;
const c = 52;
begin
  if n<=c then exit(symarr[n]);
  e:=n mod c;
  if e=0 then begin
    es:=symarr[c];
    d:=(n-c) div c;
  end
  else begin
    es:=symarr[e];
    d:=n div c;
  end;
  ds:=NumToSym(d);
  Result:=ds+es;
end;

// проверка изменений в полях типа TLength
function Changed(S,OS:TLength):boolean; overload;
begin
  if (S.Value<>OS.Value) or (S.Measure<>OS.Measure) then exit(true)
  else exit(false);
end;

function Assembledtext(node: TDOMNode): string;
var cnode: TDOMNode;
begin
   Result:='';
   cnode:= node.FirstChild;
   while assigned(cnode) do
   begin
     if cnode.NodeType=TEXT_NODE then
      Result := Result +cnode.TextContent;
     if UpperCase(cnode.NodeName)=UpperCase('text:s') then
       Result := Result+' ';
     if UpperCase(cnode.NodeName)=UpperCase('text:span') then
       Result := Result+' ';
     Result:=Result+Assembledtext(cnode);
     if UpperCase(cnode.NodeName)=UpperCase('text:p') then
       Result := Result+' ';
     cnode := cnode.NextSibling;
   end;
end;

// поиск и замена текста во всех вложенных нодах
procedure ReplaceTextInChildNodes(Node: TDOMNode; Search, Replace: string);
var
  ChildNode: TDOMNode;
  TextOfNode: String;
begin
  if Node = nil then Exit;

  ChildNode := Node.FirstChild;
  while ChildNode<>nil do
  begin
    ReplaceTextInChildNodes(ChildNode, Search, Replace);
    ChildNode := ChildNode.NextSibling;
  end;

  TextOfNode:=Node.TextContent;

  if Pos(Search,TextOfNode)>0 then
    Node.TextContent :=
                   StringReplace(TextOfNode,Search,Replace,[rfReplaceAll]);
end;

// удаление вложенных нод содержащих указанный текст
procedure RemoveNodesWithText(Node: TDOMNode; Search: string);
var
  ChildNode: TDOMNode;
  TextOfNode: String;
begin

  if Node = nil then Exit;

  ChildNode := Node.FirstChild;
  while ChildNode<>nil do
  begin
    RemoveNodesWithText(ChildNode, Search);
    ChildNode := ChildNode.NextSibling;
  end;

  TextOfNode:=Node.TextContent;

  if Pos(Search,TextOfNode)>0 then Node.ParentNode.RemoveChild(Node);

end;

{ TOdfTable }

constructor TOdfTable.Create(XMLDoc: TXMLDocument; TableName: string);
var List: TDOMNodeList;
    i:integer;
begin
  inherited Create;
  FDocument:=XMLDoc;
  List:=FDocument.DocumentElement.GetElementsByTagName('table:table');
  for i:=0 to List.Count-1 do
    if UpperCase(TDOMElement(List.Item[i]).AttribStrings['table:name'])=UpperCase(TableName)then
      begin
        FRoot:=List.Item[i];
        FName:=TableName;
        break;
      end;
  if not Assigned(FRoot) then Destroy;
end;

destructor TOdfTable.Destroy;
begin
  FDocument:=nil;
  inherited;
end;

function TOdfTable.GetRowNode(ARow: Integer): TDOMNode;
var i,j:integer;
    List: TDOMNodeList;
    Node: TDOMNode = nil;
begin
  List:=FRoot.ChildNodes;
  j:=-1;
  for i:=0 to List.Count-1 do
    begin
      if List.Item[i].NodeName='table:table-row' then
        begin
          inc(j);
          if j=ARow then
            begin
              Node:=List.Item[i];//нашли узел необходимой строки
              break;
            end;
        end;
    end;
  Result:=Node;
end;

procedure TOdfTable.MultiplyRow(ARow, Count: Integer);
var Node, Clone: TDOMNode;
    i: integer;
begin
  for i:=1 to Count do begin
    Node:=GetRowNode(ARow);
    Clone:=Node.CloneNode(true);
    FRoot.InsertBefore(Clone,Node);
  end;
end;

function TOdfTable.FindAndReplace(ARow: integer; Search, Replace: string): boolean;
begin
  try
    ReplaceTextInChildNodes(GetRowNode(ARow), Search, Replace);
    Result:=true;
  except
    Result:=false;
  end;
end;

{ TPositiveLength }

procedure TPositiveLength.SetValue(AValue: double);
begin
  if AValue <= 0 then
    raise EInOutError.Create('Value mast be > 0.')
  else
    FValue:=AValue;
end;

{ TNonNegativeLength }

procedure TNonNegativeLength.SetValue(AValue: double);
begin
  if AValue < 0 then
    raise EInOutError.Create('Value mast be >= 0.')
  else
    FValue:=AValue;
end;

{ TCellProtect }

procedure TCellProtect.SetValue(AValue: TCellProtectValues);
begin
//
end;

{ TOdtTableCell }

procedure TOdtTableCell.SetBooleanValue(AValue: boolean);
begin
  if FBooleanValue=AValue then Exit;
  FBooleanValue:=AValue;
end;

procedure TOdtTableCell.SetCurrency(AValue: string);
begin
  if FCurrency=AValue then Exit;
  FCurrency:=AValue;
end;

procedure TOdtTableCell.SetDateValue(AValue: TDateTime);
begin
  if FDateValue=AValue then Exit;
  FDateValue:=AValue;
end;

procedure TOdtTableCell.SetProtectedCell(AValue: boolean);
begin
  if FProtectedCell=AValue then Exit;
  FProtectedCell:=AValue;
end;

procedure TOdtTableCell.SetStringValue(AValue: string);
begin
  if FStringValue=AValue then Exit;
  FStringValue:=AValue;
end;

procedure TOdtTableCell.SetStyleName(AValue: string);
begin
  if FStyleName=AValue then Exit;
  FStyleName:=AValue;
end;

procedure TOdtTableCell.SetValue(AValue: double);
begin
  if FValue=AValue then Exit;
  FValue:=AValue;
end;

procedure TOdtTableCell.SetValueType(AValue: TValueTypes);
begin
  if FValueType=AValue then Exit;
  FValueType:=AValue;
end;

{ TOdtTableRowProperties }

function TOdtTableRowProperties.Equal(Obj: TOdtTableRowProperties): boolean;
begin
  Result:=false;
  if FBackgroundColor<>TOdtTableRowProperties(Obj).BackgroundColor then exit;
  if FKeepTogether<>TOdtTableRowProperties(Obj).KeepTogether then exit;
  if Changed(FMinHeight,TOdtTableRowProperties(Obj).MinHeight) then exit;
  if Changed(FHeight,TOdtTableRowProperties(Obj).Height) then exit;
  Result := inherited Equal(Obj);
end;

{ TOdtTableColumnProperties }

function TOdtTableColumnProperties.Equal(Obj: TOdtTableColumnProperties
  ): boolean;
begin
  Result:=false;
  //if FSymbol<>TOdtTableColumnProperties(Obj).Symbol then exit;
  if Changed(FWidth,TOdtTableColumnProperties(Obj).Width) then exit;
  //if Changed(FRelWidth,TOdtTableColumnProperties(Obj).RelWidth) then exit;
  Result := inherited Equal(Obj);
end;

{ TOdtTableLineProperties }

function TOdtTableLineProperties.Equal(Obj: TOdtTableLineProperties): boolean;
begin
  Result:=false;
  //if FBreakAfter<>Obj.BreakAfter then exit;
  //if FBreakBefore<>Obj.BreakBefore then exit;
  //if FUseOptimalSize<>Obj.UseOptimalSize then exit;
  Result:=true;
end;

{ TOdtTableLine }

function TOdtTableLine.Equal(Obj: TOdtTableLine): boolean;
begin
  if (FDefaultCellStyleName=Obj.DefaultCellStyleName) and
     (FStyleName=Obj.StyleName) and
     (FVisibility=Obj.Visibility) then Result:=true
  else Result:=false;
end;

{ TOdtTableProperties }

procedure TOdtTableProperties.SetWidth(w: TLength);
begin
  if w.Measure<>mCm then
    raise EInOutError.Create('Width.Measure can be only in Cm (mCm).')
  else FWidth:=w;
end;

procedure TOdtTableProperties.SetRelWidth(w: TLength);
begin
  if w.Measure<>mPercent then
    raise EInOutError.Create('RelWidth.Measure can be only in persents (mPercent).')
  else FRelWidth:=w;
end;

{ TPage }

procedure TPage.FindAndReplace(Search, Replace: string; ReplaceOnce: boolean=false);

  procedure SearchNode(Node, EndNode: TDOMNode;var aexit :boolean) ;
  var
    cNode: TDOMNode;
    txtcontent: String;
  begin
    if Node = nil then Exit;
    if aexit then Exit;
    //если нод содержит искомый текст то он разделен на чилдрены
    //где пробелы обрамлены тегом, дабы избежать повторных изменений проверим
    //чилдрен сделаем изменение и пропустим вложенные теги.

    //Внимание! Если внутри будет лежать разрыв страницы то возникнет проблема.
    //Технически при получении страницы я пофиксил разрыв переведя в корень но
    //на тот случай если перейдем к другй схеме работы со страницами оставляю предупреждение

    if Node.FirstChild<>nil then
    begin
      txtcontent:=Assembledtext(Node);
      //текст на текущей странице найден, прекращаем цикл
      if pos(UpperCase(Search),UpperCase(txtcontent))>0 then
      begin
        node.TextContent :=StringReplace(txtcontent,Search,Replace,[ rfIgnoreCase]);
        if ReplaceOnce then aexit:=true;
        Exit;
      end;
    end;

    cNode := Node.FirstChild;
    while cNode <> nil do
    begin

      SearchNode(cNode,EndNode, aexit);
      //прекращаем цыкл так как найден обрыв странцы
      if aexit then exit;
      cNode := cNode.NextSibling;
    end;

    //Ищем текст с пробелами внутри нода складывая вместе
   { if Node.FirstChild=nil then    }
      txtcontent := Node.TextContent ;
     {else
         txtcontent:=Assembledtext(Node);    }

    //текст на текущей странице найден, прекращаем цикл
    if pos(UpperCase(Search),UpperCase(txtcontent))>0 then
        begin
           node.TextContent :=StringReplace(txtcontent,Search,Replace,[ rfIgnoreCase]);
           if ReplaceOnce then aexit:=true;
        end;
    //если найден обрыв страницы прекратим поиск
    if EndNode = Node then   aexit := true;
  end;

var node,nextnode:TDOMnode;
  b: boolean;
begin
  b := false;
  //Возможно на следующих страницах будет распологатся таблица, пропустим поиск в нем
  if FNextPage=FBreakNumber then Exit;
  nextnode := FBreakNumber;
  while nextnode<>FNextPage do
  begin
      node := nextnode;
      nextnode := nextnode.NextSibling;
      SearchNode(node, FNextPage, b);
      if b then exit;
  end;

end;

function TPage.FindText(Search: UTF8string; IncludeChildren: boolean=true): UTF8string;

  procedure SearchNode(Node, EndNode: TDOMNode;var aexit :boolean) ;
  var
    cNode: TDOMNode;
    txtcontent: UTF8string;
  begin
    if Node = nil then Exit;
    if aexit then Exit;
    //если нод содержит искомый текст то он разделен на чилдрены
    //где пробелы обрамлены тегом, дабы избежать повторных изменений проверим
    //чилдрен сделаем изменение и пропустим вложенные теги.

    //Внимание! Если внутри будет лежать разрыв страницы то возникнет проблема.
    //Технически при получении страницы я пофиксил разрыв переведя в корень но
    //на тот случай если перейдем к другй схеме работы со страницами оставляю предупреждение
    if (Node.FirstChild<>nil)and IncludeChildren then
    begin
      txtcontent:=Assembledtext(Node);
      //текст на текущей странице найден, прекращаем цикл
      if UTF8pos(UTF8UpperCase(Search),UTF8UpperCase(txtcontent))>0 then
      begin
        Result := txtcontent;
        aexit:=true ;
        exit;
      end;
    end;

    cNode := Node.FirstChild;
    while cNode <> nil do
    begin

      SearchNode(cNode,EndNode, aexit);
      //прекращаем цыкл так как найден обрыв странцы
      if aexit then exit;
      cNode := cNode.NextSibling;
    end;

    //Ищем текст с пробелами внутри нода складывая вместе
   { if Node.FirstChild=nil then    }
      txtcontent := Node.TextContent ;
     {else
         txtcontent:=Assembledtext(Node);    }

    //текст на текущей странице найден, прекращаем цикл
    if UTF8pos(UTF8UpperCase(Search),UTF8UpperCase(txtcontent))>0 then
        begin
           Result := txtcontent;
           aexit := true ;
           exit;
        end;
    //если найден обрыв страницы прекратим поиск
    if EndNode = Node then   aexit := true;
  end;

var node,nextnode:TDOMnode;
  b: boolean;
begin
  b := false;
  //Возможно на следующих страницах будет распологатся таблица, пропустим поиск в нем
  if FNextPage=FBreakNumber then Exit;
  nextnode := FBreakNumber;
  while nextnode<>FNextPage do
  begin
      node := nextnode;
      nextnode := nextnode.NextSibling;
      SearchNode(node, FNextPage, b);
      if b then exit;
  end;

end;

function TPage.FindNode(Search: UTF8string; IncludeChildren: boolean=true): TDOMnode;

  procedure SearchNode(Node, EndNode: TDOMNode;var aexit :boolean) ;
  var
    cNode: TDOMNode;
    txtcontent: UTF8string;
  begin
    if Node = nil then Exit;
    if aexit then Exit;
    //если нод содержит искомый текст то он разделен на чилдрены
    //где пробелы обрамлены тегом, дабы избежать повторных изменений проверим
    //чилдрен сделаем изменение и пропустим вложенные теги.

    //Внимание! Если внутри будет лежать разрыв страницы то возникнет проблема.
    //Технически при получении страницы я пофиксил разрыв, переведя его в корень но
    //на тот случай если перейдем к другй схеме работы со страницами оставляю предупреждение

    if (Node.FirstChild<>nil)and IncludeChildren then
    begin
      txtcontent:=Assembledtext(Node);
      //текст на текущей странице найден, прекращаем цикл
      if UTF8pos(UTF8UpperCase(Search),UTF8UpperCase(txtcontent))>0 then
      begin
        Result := Node;
        aexit:=true ;
        exit;
      end;
    end;

    cNode := Node.FirstChild;
    while cNode <> nil do
    begin

      SearchNode(cNode,EndNode, aexit);
      //прекращаем цыкл так как найден обрыв странцы
      if aexit then exit;
      cNode := cNode.NextSibling;
    end;

    //Ищем текст с пробелами внутри нода складывая вместе
   { if Node.FirstChild=nil then    }
      txtcontent := Node.TextContent ;
     {else
         txtcontent:=Assembledtext(Node);    }

    //текст на текущей странице найден, прекращаем цикл
    if UTF8pos(UTF8UpperCase(Search),UTF8UpperCase(txtcontent))>0 then
        begin
           Result := Node;
           aexit := true ;
           exit;
        end;
    //если найден обрыв страницы прекратим поиск
    if EndNode = Node then   aexit := true;
  end;

var node,nextnode:TDOMnode;
  b: boolean;
begin
  Result := nil;
  b := false;
  //Возможно на следующих страницах будет распологатся таблица, пропустим поиск в нем
  if FNextPage=FBreakNumber then Exit;

  nextnode := FBreakNumber;
  while nextnode<>FNextPage do
  begin
      node := nextnode;
      nextnode := nextnode.NextSibling;
      SearchNode(node, FNextPage, b);
      if b then exit;
  end;
end;

function TPage.AppendText(aNode: TDOMNode; StyleName, Text: string): boolean;
var TextNode: TDOMNode;
    SetText: TDOMNode;
    s: string;
    i,j: integer;
begin
  try
   s := Text;
   TextNode:=FParent.Content.CreateElement('text:span');
   TDOMElement(TextNode).SetAttribute('text:style-name',StyleName);
   j := 0;
   //Проверим, нет ли пробелов перед текстом
   for i := 1 to length(s) do
      if s[i]=' ' then
        inc(j) else break;

   if j>=1 then
   begin
     AppendSpace(TextNode,j);
     delete(s,1,j);
   end;
   j:=0;
   //И тоже самое в конце
   for i := length(s) downto 1 do
      if s[i]=' ' then
        inc(j) else break;
   if j>=1 then
     delete(s,length(s)-j+1,j);

   //Добавим основной текст
   SetText:=FParent.Content.CreateTextNode(s);
   TextNode.AppendChild(SetText);

   //и пробелы в конце
   if j>=1 then AppendSpace(TextNode,j);

   aNode.AppendChild(TextNode);

   Result:=true;
  except
   Result:=false;
  end;
end;

function TPage.AppendSpace(aNode: TDOMNode; CountSpaces: integer): boolean;
var TextNode: TDOMNode;
    SetText: TDOMNode;
begin
  try
   TextNode:=FParent.Content.CreateElement('text:s');
   TDOMElement(TextNode).SetAttribute('text:c',InttoStr(CountSpaces));
   aNode.AppendChild(TextNode);
   Result:=true;
  except
   Result:=false;
  end;
end;

{procedure TPage.FindAndRemove(Search: string);
begin

end;  }

function TPage.InsertTable(Cols, Rows: integer; TableName: string): TOdtTable;
begin
  Result:=FParent.InsertTable(FNextPage.PreviousSibling,Cols, Rows, TableName);
end;

function TPage.InsertTable(InsertinBefore: TDOMnode; Cols, Rows: integer;
  TableName: string): TOdtTable;
begin
  Result:=FParent.InsertTable(InsertinBefore ,Cols, Rows, TableName);
end;

procedure TPage.DeleteTable(TableName: string);
begin
  if TableExists(TableName) then
     FParent.DeleteTable(TableName);
end;

procedure TPage.DeleteTable(aTable: TOdtTable);
begin
  if assigned(aTable)and TableExists(aTable.Name) then
     FParent.DeleteTable(aTable);

end;

function TPage.TableExists(TableName: string): boolean;
var node,nextnode:TDOMnode;
begin
  nextnode := FBreakNumber;
  //Возможно на следующих страницах будет распологатся таблица, пропустим поиск в нем
  if FNextPage=FBreakNumber then Exit;
  while nextnode<>FNextPage do
  begin
      node := nextnode;
      nextnode := nextnode.NextSibling;
      if node.NodeName='table:table' then
         if UpperCase(TDOMElement(node).AttribStrings['table:name']) = UpperCase(TableName) then
         begin
           Result := true;
           Exit;
         end;
  end;
end;

function TPage.GetTable(TableName: string): TOdtTable;
begin
  Result:=TOdtTable.Create(FParent.FContent, TableName);
  if not Assigned(Result) then
    MessageDlg('Ошибка','Таблица с названием "'+TableName+'" не найдена в документе',mtError,[mbOK],0);
end;

function TPage.GetListofTables: TStrings;
var node,nextnode:TDOMnode;
begin
  Result := TStringList.Create;
  //Возможно на следующих страницах будет распологатся таблица, пропустим поиск в нем
  if FNextPage=FBreakNumber then Exit;
  nextnode := FBreakNumber;
  while nextnode<>FNextPage.NextSibling do
  begin
      node := nextnode;
      nextnode := nextnode.NextSibling;
      if node.NodeName='table:table' then
         Result.Add(TDOMElement(node).AttribStrings['table:name']);
  end;
end;

{ TOdt }

function TOdt.StylesFindFntFace(FontName: string): TDOMNode;
var FontList: TDOMNodeList;
    Attr: TDOMNamedNodeMap;
    i,j:integer;
begin
  Result:=nil;
  FontList:=FStyles.GetElementsByTagName('style:font-face');
  for i:=0 to FontList.Length-1 do
    begin
      Attr:=FontList.Item[i].Attributes;
      for j:=0 to Attr.Length-1 do
        if Attr.Item[j].NodeName='style:name' then
          if UpperCase(Attr.Item[j].TextContent)=UpperCase(FontName)then
            begin
              Result:=FontList.Item[i];
              Exit;
            end;
    end;
end;

function TOdt.AddFont(const FontName:string;  var Doc: TXMLDocument): boolean;
var FontNode: TDOMElement;
begin
  try
    FontNode:=Doc.CreateElement('style:font-face');
    FontNode.SetAttribute('style:name',FontName);
    FontNode.SetAttribute('svg:font-family',FontName);
    Doc.GetElementsByTagName('office:font-face-decls').Item[0].AppendChild(FontNode);
    Result:=true;
  except
    Result:=false;
  end;
end;

function TOdt.AddParagraphFontStyle(StyleName: string;
                                    FontStyle: TFontStyles;
                                    FontStyleParent: boolean = false): boolean;
var Paragraph: TDOMElement;
begin
  try
   Paragraph:=FContent.CreateElement('style:style');
   Paragraph.SetAttribute('style:name',StyleName);
   Paragraph.SetAttribute('style:family','text');
   //добавляем новый узел параграфа
   FContent.GetElementsByTagName('office:automatic-styles').Item[0].AppendChild(Paragraph);
   Paragraph:=FContent.CreateElement('style:text-properties');

   if ftItalic in FontStyle then
     begin
       Paragraph.SetAttribute('fo:font-style','italic');
       Paragraph.SetAttribute('style:font-style-asian','italic');
       Paragraph.SetAttribute('style:font-style-complex','italic');
     end;
   if ftBold in FontStyle then
     begin
       Paragraph.SetAttribute('fo:font-weight','bold');
       Paragraph.SetAttribute('style:font-weight-asian','bold');
       Paragraph.SetAttribute('style:font-weight-complex','bold');
     end;
   if ftUnderline in FontStyle then
     begin
       Paragraph.SetAttribute('style:text-underline-style','solid');
       Paragraph.SetAttribute('style:text-underline-width','auto');
       Paragraph.SetAttribute('style:text-underline-color','font-color');
     end;
   if not(ftItalic in FontStyle) then
     begin
       Paragraph.SetAttribute('fo:font-style','normal');
       Paragraph.SetAttribute('style:font-style-asian','normal');
       Paragraph.SetAttribute('style:font-style-complex','normal');
     end;
   if not(ftBold in FontStyle) then
     begin
       Paragraph.SetAttribute('fo:font-weight','normal');
       Paragraph.SetAttribute('style:font-weight-asian','normal');
       Paragraph.SetAttribute('style:font-weight-complex','normal');
     end;
   if not(ftUnderline in FontStyle) then
     begin
       Paragraph.SetAttribute('style:text-underline-style','normal');
       Paragraph.SetAttribute('style:text-underline-width','normal');
       Paragraph.SetAttribute('style:text-underline-color','normal');
     end;

    //добавляем описание шрифта
    FContent.GetElementsByTagName('style:style').Item[FContent.GetElementsByTagName('style:style').Length-1].AppendChild(Paragraph);
    Result:=true;
  except
    Result:=false;
  end;
end;

function TOdt.AddParagraphFontStyle(FontStyle: TFontStyles;
  FontStyleParent: boolean): string;
begin
  Result := SearchFreeStyleName('P1');
  if Result='' then exit;
  AddParagraphFontStyle(Result,FontStyle,FontStyleParent);
end;

function TOdt.GetParagraphFontStyle(FontStyle: TFontStyles): string;
var
  Paragraph: TDOMElement;

  function CountStyles():integer;
  var
  Element: TFontStyle;
  begin
    Result:= 0;
    for Element := Low(TFontStyle) to High(TFontStyle) do
      if Element in FontStyle then
        Inc(Result);
  end;
  function ItalicExist(): boolean;
  begin
    Result := (
              (LowerCase(Paragraph.AttribStrings['fo:font-style'])='italic') and
              (LowerCase(Paragraph.AttribStrings['style:font-style-asian'])='italic') and
              (LowerCase(Paragraph.AttribStrings['style:font-style-complex'])='italic'))
  end;
  function BoldExist(): boolean;
  begin
    Result := ((LowerCase(Paragraph.AttribStrings['fo:font-weight'])='bold') and
               (LowerCase(Paragraph.AttribStrings['style:font-weight-asian'])='bold') and
               (LowerCase(Paragraph.AttribStrings['style:font-weight-complex'])='bold'))
  end;
  function UnderlineExist(): boolean;
  begin
    Result := ((LowerCase(Paragraph.AttribStrings['style:text-underline-style'])='solid')and
               (LowerCase(Paragraph.AttribStrings['style:text-underline-width'])='auto')and
               (LowerCase(Paragraph.AttribStrings['style:text-underline-color'])='font-color'))
  end;
  //Для того что бы не зависить от стиля родителя
  //в стиль добавляются другие параметры выставляемые как нормал
  function ItalicNormalExist(): boolean;
  begin
    Result:= not Paragraph.hasAttribute('fo:font-style');
    if Result then Exit;
    Result := ((LowerCase(Paragraph.AttribStrings['fo:font-style'])='normal') and
              (LowerCase(Paragraph.AttribStrings['style:font-style-asian'])='normal') and
              (LowerCase(Paragraph.AttribStrings['style:font-style-complex'])='normal'))
  end;
  function BoldNormalExist(): boolean;
  begin
    Result:= not Paragraph.hasAttribute('fo:font-weight');
    if Result then Exit;
    Result := ((LowerCase(Paragraph.AttribStrings['fo:font-weight'])='normal') and
               (LowerCase(Paragraph.AttribStrings['style:font-weight-asian'])='normal') and
               (LowerCase(Paragraph.AttribStrings['style:font-weight-complex'])='normal'))
  end;
  function UnderlineNormalExist(): boolean;
  begin
    Result:= not Paragraph.hasAttribute('style:text-underline-style');
    if Result then Exit;
    Result := ((LowerCase(Paragraph.AttribStrings['style:text-underline-style'])='normal') and
              (LowerCase(Paragraph.AttribStrings['style:text-underline-width'])='normal') and
              (LowerCase(Paragraph.AttribStrings['style:text-underline-color'])='normal'))
  end;
var
    root,node :TDOMnode;
    List: TDOMnodeList;
    i:integer;
    Sum,iItalic,iBold,iUnderline: boolean;
    fCountStyles: integer;
begin
  //Это мы так подсчитываем число элементов,
  //ну что бы в стиль входили только эти элементы
  // (Paragraph.Attributes.Length=3*fCountStyles);
    Result := '';
    fCountStyles := CountStyles;
    root := Content.GetElementsByTagName('office:automatic-styles').Item[0];
    list := root.GetChildNodes;
    //Проверим есть ли такой же стиль
    for i:=0 to list.Count-1 do
        begin

          node:= list[i].FirstChild;

          if node=nil then continue;
          if node.NodeName<>'style:text-properties' then continue  ;

          Paragraph := TDOMElement(node);
          //если ненайдем пропустим
          if Paragraph.NodeName<>'style:text-properties' then continue;
          { and
                 (Paragraph.Attributes.Length=3*fCountStyles)};

          if ftItalic in FontStyle then
            iItalic:= ItalicExist
               else iItalic:= ItalicNormalExist;

          if ftBold in FontStyle then
            iBold:= BoldExist
              else iBold:= BoldNormalExist;

          if ftUnderline in FontStyle then
              iUnderline:= UnderlineExist
              else  iUnderline:= UnderlineNormalExist ;

          if not (iItalic and iBold and iUnderline)   then continue;
          Result :=  TDOMElement(List[i]).AttribStrings['style:name'];
          Exit;
        end;
end;

function TOdt.AddParagraphTextPosition(StyleName: string; TextPosition: TTextPosition
  ): boolean;
var Paragraph: TDOMElement;
begin
  try
   Paragraph:=FContent.CreateElement('style:style');
   Paragraph.SetAttribute('style:name',StyleName);
   Paragraph.SetAttribute('style:family','text');
   //добавляем новый узел параграфа
   FContent.GetElementsByTagName('office:automatic-styles').Item[0].AppendChild(Paragraph);
   Paragraph:=FContent.CreateElement('style:text-properties');

   case TextPosition of
     tpCenter:begin
                Paragraph:=FContent.CreateElement('style:paragraph-properties');
                Paragraph.SetAttribute('fo:text-align','center');
                Paragraph.SetAttribute('style:justify-single-word','false');
              end;
     tpRight:begin
               Paragraph:=FContent.CreateElement('style:paragraph-properties');
               Paragraph.SetAttribute('fo:text-align','end');
               Paragraph.SetAttribute('style:justify-single-word','false');
             end;
     tpJustify:begin
                 Paragraph:=FContent.CreateElement('style:paragraph-properties');
                 Paragraph.SetAttribute('fo:text-align','justify');
                 Paragraph.SetAttribute('style:justify-single-word','false');
               end;
   end;
   //добавляем описание положения текста
   FContent.GetElementsByTagName('style:style').Item[FContent.GetElementsByTagName('style:style').Length-1].AppendChild(Paragraph);
   Result:=true;
  except
    Result:=false;
  end;
end;

function TOdt.AddParagraphFontSize(StyleName: string; FontSize: integer
  ): boolean;
var
  Paragraph: TDOMElement;
begin
  try
   Paragraph:=FContent.CreateElement('style:style');
   Paragraph.SetAttribute('style:name',StyleName);
   Paragraph.SetAttribute('style:family','text');
   //добавляем новый узел параграфа
   FContent.GetElementsByTagName('office:automatic-styles').Item[0].AppendChild(Paragraph);

   Paragraph:=FContent.CreateElement('style:text-properties');
   Paragraph.SetAttribute('fo:font-size',IntToStr(FontSize)+'pt');
   Paragraph.SetAttribute('style:font-size-asian',IntToStr(FontSize)+'pt');
   Paragraph.SetAttribute('style:font-size-complex',IntToStr(FontSize)+'pt');
   FContent.GetElementsByTagName('style:style').Item[FContent.GetElementsByTagName('style:style').Length-1].AppendChild(Paragraph);
 except
   Result:=false;
 end;
end;

function TOdt.AddParagraphFontSize(FontSize: integer): string;
begin
  Result := SearchFreeStyleName('A1');
  if Result='' then exit;
  AddParagraphFontSize(Result, FontSize);
end;

function TOdt.AddParagraphStyle(StyleName, FontName: string;
  FontSize: integer; FontStyle: TFontStyles; TextPosition: TTextPosition): boolean;
var Paragraph: TDOMElement;
begin
try
 Paragraph:=FContent.CreateElement('style:style');
 Paragraph.SetAttribute('style:name',StyleName);
 Paragraph.SetAttribute('style:family','paragraph');
 Paragraph.SetAttribute('style:parent-style-name','Standard');
 //добавляем новый узел параграфа
 FContent.GetElementsByTagName('office:automatic-styles').Item[0].AppendChild(Paragraph);

 Paragraph:=FContent.CreateElement('style:text-properties');
 Paragraph.SetAttribute('fo:font-size',IntToStr(FontSize)+'pt');
 Paragraph.SetAttribute('style:font-name',FontName);
 if ftItalic in FontStyle then
   begin
     Paragraph.SetAttribute('fo:font-style','italic');
     Paragraph.SetAttribute('style:font-style-asian','italic');
     Paragraph.SetAttribute('style:font-style-complex','italic');
   end;
 if ftBold in FontStyle then
   begin
     Paragraph.SetAttribute('fo:font-weight','bold');
     Paragraph.SetAttribute('style:font-weight-asian','bold');
     Paragraph.SetAttribute('style:font-weight-complex','bold');
   end;
 if ftUnderline in FontStyle then
   begin
     Paragraph.SetAttribute('style:text-underline-style','solid');
     Paragraph.SetAttribute('style:text-underline-width','auto');
     Paragraph.SetAttribute('style:text-underline-color','font-color');
   end;

  //добавляем описание шрифта
  FContent.GetElementsByTagName('style:style').Item[FContent.GetElementsByTagName('style:style').Length-1].AppendChild(Paragraph);

  case TextPosition of
    tpCenter:begin
               Paragraph:=FContent.CreateElement('style:paragraph-properties');
               Paragraph.SetAttribute('fo:text-align','center');
               Paragraph.SetAttribute('style:justify-single-word','false');
             end;
    tpRight:begin
              Paragraph:=FContent.CreateElement('style:paragraph-properties');
              Paragraph.SetAttribute('fo:text-align','end');
              Paragraph.SetAttribute('style:justify-single-word','false');
            end;
    tpJustify:begin
                Paragraph:=FContent.CreateElement('style:paragraph-properties');
                Paragraph.SetAttribute('fo:text-align','justify');
                Paragraph.SetAttribute('style:justify-single-word','false');
              end;
  end;
  //добавляем описание положения текста
  FContent.GetElementsByTagName('style:style').Item[FContent.GetElementsByTagName('style:style').Length-1].AppendChild(Paragraph);
  Result:=true;
  except
    Result:=false;
  end;
end;

function TOdt.AppendText(StyleName, Text: string): boolean;
var TextNode: TDOMNode;
    SetText: TDOMNode;
begin
  try
   TextNode:=FContent.CreateElement('text:p');
   TDOMElement(TextNode).SetAttribute('text:style-name',StyleName);
   SetText:=FContent.CreateTextNode(Text);
   TextNode.AppendChild(SetText);
   FContent.GetElementsByTagName('office:text').Item[FContent.GetElementsByTagName('office:text').Length-1].AppendChild(TextNode);
   Result:=true;
  except
   Result:=false;
  end;
end;

function TOdt.isStyleExists( text: string ): boolean;
var
    root:TDOMnode;
    List: TDOMnodeList;
    i:integer;
begin
    Result := false;
    root := Content.GetElementsByTagName('office:automatic-styles').Item[0];
    list := root.GetChildNodes;
    //Проверим есть ли такой же стиль
    for i:=0 to list.Count-1 do
      if TDOMElement(list[i]).AttribStrings['style:name']=text then
        begin
          Result := true;
          exit;
        end;
end;

function TOdt.SearchFreeStyleName(StyleName: string; DontSearchinTable: Boolean=true): string;
  function isDigit( c: Char ): boolean;
  begin
     Result := c in ['0'..'9']
  end;

  function SymtoInt( c: Char ): Integer;
  var i: integer;
  begin
    Result := 1;
    for i:=1 to 52 do
        if c = symarr[i] then Result := i;

  end;

  var
      Sybols, Digit,newStyleName: string;
      i,m,n:integer;
begin
    if DontSearchinTable then
    begin
      if pos('.',StyleName)>0 then
        delete(StyleName,pos('.',StyleName),length(StyleName));
    end
      else
      {Поиск стиля внутри таблицы еще не создан}
        Exit;

    // такой стиль найден.
    if isStyleExists(StyleName) then
    begin

        newStyleName:= StyleName;
        //разобьем текст на цифры и символы
        Sybols:='';
        Digit:='';
        for i:=1 to length(newStyleName) do
        begin
        if not isDigit(newStyleName[i]) then
           Sybols:= Sybols+ newStyleName[i];
        if isDigit(newStyleName[i]) then
           Digit:= Digit+ newStyleName[i];
        end;
        //Название стиля таблиц имеет большую длинну и ее моно увличиить.
        //например "Таблица1"
        if length(Sybols)>2 then
        begin
           //сгенерируем новое имя стиля и проверим существует ли он
           i:=strtoint(Digit);
           repeat
             inc(i);
           until not isStyleExists(Sybols+inttostr(i));
           Result:= Sybols+inttostr(i);
           exit;
        end;
        //картинки и другие рисуемые обьекты имеют длинну в 2 маленьких символа.
        //например "fr1"
        if length(Sybols)=2 then
        begin
           //на всякий пожарный уменьшим символы если он каким то макаром будет большими
           Sybols:=LowerCase(Sybols);
           i := SymtoInt(Sybols[1]);
           m := SymtoInt(Sybols[2]);
           n := strtoint(Digit);
           //Сложная вещь надеюсь багов не будет!
           //Суть 3х вложенных циклов в том что бы пробежатся по таблице маленьких символов
           //начиная с текущей, и выбрать уникальное имя
           //Цифры подбираются в диапазоне от 1 до 9и
           while true do
           begin
             if i >= 52 then i := 27;

             while true do
             begin
               if m >= 52 then m := 27;

               while true do
               begin
                 if n >= 9 then n := 1;
                   if not  isStyleExists(symarr[i] + symarr[m] + inttostr(n)) then
                   begin
                     Result := symarr[i] + symarr[m] + inttostr(n);
                     Exit;
                   end;
                 inc(n);
               end;
               inc(m);
             end;
             inc(i);
           end;
        end;
        //Стили текста обрамляются одним большим символом.
        //P23
        if length(Sybols)=1 then
        begin
          //на всякий пожарный увличим символ если он каким то макаром будет маленьким
            Sybols:=UpperCase(Sybols);
            i := SymtoInt(Sybols[1]);
            n := strtoint(Digit);
           //Сложная вещь надеюсь багов не будет!
           //Суть 2х вложенных циклов в том что бы пробежатся по таблице больших символов
           //начиная с текущей, и выбрать уникальное имя
           //Цифры подбираются в диапазоне от 1 до 99и
           while true do
           begin
             if i >= 27 then i := 1;

               while true do
               begin
                 if n >= 99 then n := 1;
                   if not  isStyleExists(symarr[i] + inttostr(n)) then
                   begin
                     Result := symarr[i] + inttostr(n);
                     Exit;
                   end;
                 inc(n);
               end;
             inc(i);
           end;
        end;
        Result:=Sybols;
        Exit;
      end
      else
      //ненайден, возвратим тот же стиль, не нилы же кидать :)
      Result:= StyleName;
end;

procedure TOdt.AddPage(aIndex: integer);
var node,addel,addch: TDOMNode;
begin
end;

procedure TOdt.CashPages();
 var List: TDOMNodeList;
     Node: TDOMNode;
     i,j:integer;
     stylebreaker:array of string;

   Procedure AddPageinArray(aNode: TDOMNode;onlyone: boolean = false);
   var l: integer;
     List: TDOMNodeList;
   begin
       //добавляем первую страницу если их 2
       if length(FCashPages) = 0 then
       begin
         List:= FContent.GetElementsByTagName('office:text').Item[0].ChildNodes;
         l := Length(FCashPages);
         SetLength(FCashPages, l+1);
         if  UpperCase(TDOMElement(List.Item[0]).TagName)<>UpperCase('office:forms') then
            FCashPages[l].BreakNumber :=  List.Item[0]
            else
            if  UpperCase(TDOMElement(List.Item[1]).TagName)<>UpperCase('text:sequence-decls') then
               FCashPages[l].BreakNumber :=  List.Item[1]
               else
                 FCashPages[l].BreakNumber :=  List.Item[2];
       end;
       //Страниц несколько, само собой
       if not onlyone then
       if length(FCashPages) <> 0 then
       begin
         l := Length(FCashPages);
         SetLength(FCashPages, l+1);
         FCashPages[l].BreakNumber := aNode;
       end;

   end;

   procedure SearchBreak(aNode: TDOMNode);
   var i,j: Integer;
       nodebreak: string;
       childnode:TDOMNodeList;
   begin
       If (aNode=nil) then exit;
       childnode := aNode.ChildNodes;

       for i:=0 to childnode.Count-1 do
       begin

         if UpperCase(childnode[i].NodeName)=UpperCase('text:soft-page-break') then
           AddPageinArray(childnode[i].ParentNode);

         //Отсеиваем элементы не являющиеся узлами такие как простой текст.
         if childnode[i].HasAttributes then
         begin
           nodebreak:=TDOMElement(childnode[i]).AttribStrings['text:style-name'];
           //Ищем разрывы страницы внутри стиля
           if UpperCase(nodebreak) <>UpperCase('Standard') then
           begin

             for j:=0 to Length(stylebreaker)-1 do
                 if UpperCase(stylebreaker[j]) = UpperCase(nodebreak) then
                 begin
                   AddPageinArray(childnode[i]);
                 end;

           end;
         end;
         SearchBreak(childnode[i]);
       end;
   end;

 begin
  SetLength(FCashPages, 0);
  j:=0;
  //Кэшируем стили с разрывами страниц
  List:=FContent.GetElementsByTagName('style:paragraph-properties');
  for i:=0 to List.Count-1 do
  begin
         if UpperCase(TDOMElement(List.Item[i]).AttribStrings['fo:break-before'])=UpperCase('page') then
         begin
           inc(j);
           SetLength(stylebreaker,j);
           stylebreaker[j-1]:=TDOMElement(List.Item[i].ParentNode).AttribStrings['style:name'];
         end;
  end;
  node:= FContent.GetElementsByTagName('office:text').Item[0];
  SearchBreak(node);
  //Страница только одна, других больше нету
  if (CountPages=0) then AddPageinArray(nil, true)
end;

procedure TOdt.RemovePage(aIndex: integer; WithBreakPageonTable : boolean = true);
begin
  RemovePage(aIndex, aIndex, WithBreakPageonTable);
end;

procedure TOdt.RemovePage(aFromIndex, aToIndex: integer;
  WithBreakPageonTable: boolean);
var i,endpage:integer ;
    list:TDOMnodeList;
    root,cnode,endnode,node:TDOMnode;
begin
  root := FContent.GetElementsByTagName('office:text').Item[0];
  list := root.GetChildNodes;

  //если страница последняя, завершение цикла ставим последний узел
  if aToIndex=CountPages-1 then
    endnode := GetEndPages.PreviousSibling
    else
       endnode := FCashPages[aToIndex+1].BreakNumber;

  node:= FCashPages[aFromIndex].BreakNumber;

  //Если разрыв страницы находится в таблице то удаляем вместе с таблицей
  if WithBreakPageonTable then
  begin
     if node.ParentNode<>root then
     begin
        node := GetFixParentNode(node);
     end;
     if endnode.ParentNode<>root then
     begin
        endnode := GetFixParentNode(endnode);
        endnode := endnode.NextSibling;
     end;
  end else
      //иначе удаляем без таблицы
      begin
         if node.ParentNode<>root then
         begin
            node := GetFixParentNode(node);
            node := node.NextSibling;
         end;
         if endnode.ParentNode<>root then
         begin
            endnode := GetFixParentNode(endnode);
            endnode := endnode.PreviousSibling;
         end;
      end;

 while node<>endnode do
  begin
       cnode:=node;
       node:=node.NextSibling;
       cnode.ParentNode.RemoveChild(cnode);
  end;
  CashPages();
end;

procedure TOdt.MovePage(Source, Target: integer; WithBreakPageonTable : boolean = true);
begin
  if Source =  Target then exit;
  MovePage(Source, Source, Target, WithBreakPageonTable);
end;

procedure TOdt.MovePage(aFromSource, aToSource, Target: integer;
  WithBreakPageonTable: boolean);
var
    node,endnode,targetnode,root,nextnode: TDOMnode;
        list:TDOMnodeList;
begin
  //если цель находится внутри перемещаемых обьектов то выходим
  if (aFromSource >=  Target) and (aToSource <=  Target) then exit;

  root := FContent.GetElementsByTagName('office:text').Item[0];
  list := root.GetChildNodes;

//------------ Источник -----------------------------

    //Конец исходной страницы
    node := FCashPages[aFromSource].BreakNumber;
    //Начало исходной страницы
    if aToSource = CountPages-1 then
       endnode := List[List.Count-1]
      else
        endnode:= FCashPages[aToSource+1].BreakNumber;

    //Если разрыв страницы находится в таблице то удаляем вместе с таблицей
    if WithBreakPageonTable then
    begin
       if node.ParentNode<>root then
       begin
          node := GetFixParentNode(node);
       end;
       if endnode.ParentNode<>root then
       begin
          endnode := GetFixParentNode(endnode);
          endnode := endnode.NextSibling;
       end;
    end else
        //иначе удаляем без таблицы
        begin
           if node.ParentNode<>root then
           begin
              node := GetFixParentNode(node);
              node := node.NextSibling;
           end;
           if endnode.ParentNode<>root then
           begin
              endnode := GetFixParentNode(endnode);
              endnode := endnode.PreviousSibling;
           end;
        end;

//------------------- Цель -----------------------------

    //Начало цели
    if Target = 0 then
            targetnode :=  GetBeginPages
      else
      //Цель последняя страница
      if Target = CountPages-1 then
         targetnode := List[List.Count-1]
        else
        //назначение отличаются
          if (aFromSource < Target)and (aToSource < Target) then
             begin
                targetnode:= FCashPages[Target+1].BreakNumber;
                //проверим узел находится в таблице?
                if targetnode.ParentNode<>root then
                  targetnode := GetFixParentNode(targetnode);
             end
             else
               if (aFromSource > Target)and (aToSource > Target) then
               begin
                  targetnode := FCashPages[Target+1].BreakNumber.PreviousSibling;
                  //проверим узел находится в таблице?
                  if targetnode.ParentNode<>root then
                    targetnode := GetFixParentNode(targetnode);
               end;

      while node<>endnode do
      begin
        nextnode := node;
        node := node.NextSibling;
        nextnode.ParentNode.InsertBefore(nextnode, targetnode);
      end;
     CashPages();
end;

function TOdt.FindPage(Search: string): Integer;
var
  i: integer;
  fpage: TPage;
begin
  Result:=-1;
  //пробегаемся по массиву страниц
  for i:=0 to CountPages-1 do
  begin
    fpage := page[i];
    if fpage.FindText(Search)<>'' then
    begin
      Result:=i;
      break;
    end;
    fpage.Destroy;
  end;
end;

procedure TOdt.Copy(Source, Target: Integer; WithBreakPageonTable : boolean = true);
begin
  Copy(Source, Source, Target, WithBreakPageonTable);
end;

procedure TOdt.Copy(aFromSource, aToSource, Target: Integer;
  WithBreakPageonTable: boolean);
var
    node,endnode,targetnode,root,clonenode,nextnode: TDOMnode;
        list:TDOMnodeList;
begin
//если цель находится внутри перемещаемых обьектов то выходим
//if (aFromSource >=  Target) and (aToSource <=  Target) then exit;

  root := FContent.GetElementsByTagName('office:text').Item[0];
  list := root.GetChildNodes;
//--------------------Источник------------------------------

    //Конец исходной страницы
    node := FCashPages[aFromSource].BreakNumber;
    //Начало исходной страницы
    if aToSource = CountPages-1 then
       //Если указать и цель и источник одну страницу то выйдет беспокнечное копирование
    //Так что укажем сдесь копировать только до предпоследней строки страницы
       endnode := GetEndPages.PreviousSibling
      else
        endnode:= FCashPages[aToSource+1].BreakNumber;

    //Если разрыв страницы находится в таблице то копируем с таблицей
    if WithBreakPageonTable then
    begin
       if node.ParentNode<>root then
       begin
          node := GetFixParentNode(node);
       end;
       if endnode.ParentNode<>root then
       begin
          endnode := GetFixParentNode(endnode);
          endnode := endnode.NextSibling;
       end;
    end else
        //иначе копируем без таблицей
        begin
           if node.ParentNode<>root then
           begin
              node := GetFixParentNode(node);
              node := node.NextSibling;
           end;
           if endnode.ParentNode<>root then
           begin
              endnode := GetFixParentNode(endnode);
              endnode := endnode.PreviousSibling;
           end;
        end;

//--------------------Цель------------------------------

        //Начало цели
        if Target = 0 then
                targetnode :=  GetBeginPages
          else
          //Цель последняя страница
          if Target = CountPages-1 then
             targetnode := GetEndPages;
            //else
            //назначение отличаются
              if (aFromSource < Target)and (aToSource < Target) then
                 begin
                    targetnode:= FCashPages[Target+1].BreakNumber;
                    //проверим узел находится в таблице?
                    if targetnode.ParentNode<>root then
                      targetnode := GetFixParentNode(targetnode);
                 end
                 else
                   if (aFromSource > Target)and (aToSource > Target) then
                   begin
                      targetnode := FCashPages[Target+1].BreakNumber.PreviousSibling;
                      //проверим узел находится в таблице?
                      if targetnode.ParentNode<>root then
                        targetnode := GetFixParentNode(targetnode);
                   end;

//-------------------Action --------------------------------
      while node<>endnode do
      begin
        nextnode := node;
        node := node.NextSibling;
        clonenode := nextnode.CloneNode(true);
        //Переименуем таблицу
        if (clonenode.NodeName)='table:table' then
          TDOMElement(clonenode).AttribStrings['table:name'] := GetNewTableName;
        nextnode.ParentNode.InsertBefore(clonenode, targetnode);
      end;
      //Докопируем посленюю строку, ибо она не копируется если идет клонирование страницы.
      //Прямое копирование вызывает Безконечный цыкл, так что копируем до последней строки
      clonenode:=nextnode.NextSibling.CloneNode(true);
      if (clonenode.NodeName)='table:table' then
        TDOMElement(clonenode).AttribStrings['table:name'] := GetNewTableName;
      nextnode.ParentNode.InsertBefore(clonenode, targetnode);

     CashPages();
end;

//Копирование страницы из одного документа в другой
//по мойму самый геморойный код, нуно перенести не ток даные но и стили и шрифты
procedure TOdt.Copy(SourceDocument: TOdt; Source, Target: Integer; WithBreakPageonTable : boolean = true);

var
    node, endnode, importednode, beginnode, targetnode, root, nextnode: TDOMnode;
    list:TDOMnodeList;
    fpage,fnextpage: TPage;
    farraystyles: array of string;
    farraynewstyles: array of string;

    Sourcestyleroot:TDOMnode;
    SourcestyleList: TDOMnodeList;
    Targetstyleroot:TDOMnode;
    TargetstyleList: TDOMnodeList;

  //Добавим название стиля, если конечно он еще там не добавлен
  procedure AddStyletoCash(style: string);
  var i,L:integer;
  begin
     for i:=0 to length(farraystyles)-1 do
       if farraystyles[i]=style then exit;

       L:= length(farraystyles);
       Setlength(farraystyles,L+1);
       farraystyles[L] := style;
  end;

  //поищем все стили которые встречаются в структуре
  procedure CashStyles(Node: TDOMNode);
  var
    cNode: TDOMNode;
    nodename: String;
  begin
    if Node = nil then Exit;

    cNode := Node.FirstChild;
    while cNode <> nil do
    begin
      CashStyles(cNode);
      cNode := cNode.NextSibling;
    end;

    //ищем стиль в тексте, таблице, рисунке
    //если какой то стиль не перенесся то добавте сюда
    if Node.NodeType<>TEXT_NODE then
    begin
      if TDOMElement(Node).hasAttribute('text:style-name') then
         nodename:= TDOMElement(Node).AttribStrings['text:style-name'];
      if TDOMElement(Node).hasAttribute('table:style-name') then
         nodename:= TDOMElement(Node).AttribStrings['table:style-name'];
      if TDOMElement(Node).hasAttribute('draw:style-name') then
         nodename:= TDOMElement(Node).AttribStrings['draw:style-name'];
      AddStyletoCash(nodename);
    end;

  end;
  function SearchStyle(style: string): TDOMnode;
  var
      i:integer;
  begin
      Result := nil;
      for i:=0 to Sourcestylelist.Count-1 do
        if TDOMElement(Sourcestylelist[i]).AttribStrings['style:name']=style then
          begin
            Result := Sourcestylelist[i];
            exit;
          end;
  end;

  function ImportStyle(node: TDOMnode): TDOMnode;
  begin
    Result:= FContent.ImportNode(node, true);
    //targetnode.ParentNode.InsertBefore(FContent.ImportNode(nextnode,true), targetnode);
    targetstyleroot.InsertBefore(Result, targetstyleroot.LastChild);
  end;

  procedure SynchronizationStyles();
  var i, j, z: integer;
      root, node: TDOMnode;
      List: TDOMnodeList;

      stylename,newstylename: string;
  begin
    //перепишем данные в новый массив
    SetLength(farraynewstyles,Length(farraystyles));

    for i:=0 to Length(farraystyles)-1 do
    begin
       farraynewstyles[i]:=farraystyles[i];
    end;

    //Синхронизируем стили
    for i:=0 to Length(farraynewstyles)-1 do
    begin
   //   if farraynewstyles[i]='' then continue;
        // form1.Memo1.Lines.Add(farraynewstyles[i]);

       //переносим обычные стили
       if pos('Таблица',farraynewstyles[i])=0 then
       begin
         node := SearchStyle(farraynewstyles[i]);

         if assigned(node) then
         begin
           //переименуем если такой стиль сущществует
           if isStyleExists(farraynewstyles[i]) then
           begin
             farraynewstyles[i] := SearchFreeStyleName(farraynewstyles[i]);
             TDOMElement(ImportStyle(node)).AttribStrings['style:name'] := farraynewstyles[i];
           end
             else
               ImportStyle(node);
         end;
       end;

       //отсеиваем таблицы со стилями колонок и сталбцов
       if (pos('Таблица',farraynewstyles[i])>0) and (pos('.',farraynewstyles[i])>0) then
       continue;

       //Найден корневой стиль таблицы
       if (pos('Таблица',farraynewstyles[i])>0) and (pos('.',farraynewstyles[i])=0) then
       begin
          stylename := farraynewstyles[i];
          //переименуем таблицу если такая сущществует
          if isStyleExists(farraynewstyles[i]) then
            newstylename := SearchFreeStyleName(farraynewstyles[i])
          else newstylename := stylename;


          for j:=0 to Length(farraynewstyles)-1 do
            //Найдем все стили таблицы для переименования
            if pos(stylename, farraynewstyles[j])>0 then
            begin
              node := SearchStyle(farraynewstyles[j]);
              //переименуем название таблицы для внутренних стилей таблицы "Таблица1.А1"
              farraynewstyles[j]:= StringReplace(farraynewstyles[j], stylename, newstylename,[]);

              if assigned(node) then
              begin
                node := ImportStyle(node);
                TDOMElement(node).AttribStrings['style:name'] := farraynewstyles[j];
              end;
            end;
       end;
    end;
  end;

  {
  ЗДЕСЬ Лучше всего реализовать механизм импорта картинок!
  }
  procedure SynchronizationData(node: TDOMnode);
  var
    cNode: TDOMNode;
    nodename: String;
    i: integer;
  begin
    if Node = nil then Exit;

    cNode := Node.FirstChild;
    while cNode <> nil do
    begin
      SynchronizationData(cNode);
      cNode := cNode.NextSibling;
    end;
    //ищем стиль который ранее зарегестрирован,
    //если такой есть то переделаем его на новый из массива переделанных стилей
    for i:=0 to length(farraystyles)-1 do
    begin
       if TDOMElement(Node).hasAttribute('text:style-name') then
          if farraystyles[i] = TDOMElement(Node).AttribStrings['text:style-name'] then
             TDOMElement(Node).AttribStrings['text:style-name'] := farraynewstyles[i];
       if TDOMElement(Node).hasAttribute('table:style-name') then
          if farraystyles[i] = TDOMElement(Node).AttribStrings['table:style-name'] then
             TDOMElement(Node).AttribStrings['table:style-name'] := farraynewstyles[i];
       if TDOMElement(Node).hasAttribute('draw:style-name') then
          if farraystyles[i] = TDOMElement(Node).AttribStrings['draw:style-name'] then
             TDOMElement(Node).AttribStrings['draw:style-name'] := farraynewstyles[i];
    end;
  end;
  //Синхронизация шрифтов
  procedure SynchronizationFont();
  var farryfonts: array of string;
      farrynewfonts: array of string;
    procedure AddFonttoCash(style: string);
    var i,L:integer;
    begin
       for i:=0 to length(farryfonts)-1 do
         if farryfonts[i]=style then exit;

         L:= length(farryfonts);
         Setlength(farryfonts,L+1);
         farryfonts[L] := style;
    end;
  var
    Sourcefontroot:TDOMnode;
    SourcefontList: TDOMnodeList;
    Targetfontroot:TDOMnode;
    TargetfontList: TDOMnodeList;
    i: integer;
  begin
      //рут дерева стилей в источнике и в пункте назначения
      sourcestyleroot := SourceDocument.Content.GetElementsByTagName('office:font-face-decls').Item[0];
      sourcestylelist := sourcestyleroot.GetChildNodes;
      targetstyleroot := Content.GetElementsByTagName('office:font-face-decls').Item[0];
      targetstylelist := targetstyleroot.GetChildNodes;

      for i:=0 to length(farraystyles) do
      begin
    //    node := SearchStyle(farraynewstyles[j]);
      end;
      {НЕ ЗАКОНЧЕНО
      Идея в том что бы записать все шрифты которые встречаются в стилях затем
      сравнить и переименовать\оставить так как есть и переместить шрифт    }
  end;

begin


  fpage := SourceDocument.Page[Source];
  if not assigned(fpage)and not assigned(SourceDocument) then exit;

  root := SourceDocument.Content.GetElementsByTagName('office:text').Item[0];
  list := root.GetChildNodes;

  //рут дерева стилей в источнике и в пункте назначения
  sourcestyleroot := SourceDocument.Content.GetElementsByTagName('office:automatic-styles').Item[0];
  sourcestylelist := sourcestyleroot.GetChildNodes;
  targetstyleroot := Content.GetElementsByTagName('office:automatic-styles').Item[0];
  targetstylelist := targetstyleroot.GetChildNodes;

    //Конец исходной страницы
    node := fpage.BreakNumber;
    //Начало исходной страницы
    if Source+1 = SourceDocument.CountPages-1 then
       endnode := List[List.Count-1]
      else
      begin
        fnextpage := SourceDocument.Page[Source+1];
        endnode:= fnextpage.BreakNumber;
      end;

    //Начало цели
    if Target = 0 then
            targetnode :=  GetBeginPages()
      else
      //смотря как копируем в прямом или обратном порядке
      if source < Target then
            targetnode:= FCashPages[Target+1].BreakNumber
         else
           if source > Target then
              targetnode := FCashPages[Target+1].BreakNumber.PreviousSibling;

    //Если разрыв страницы находится в таблице то копируем с таблицей
    if WithBreakPageonTable then
    begin
       if node.ParentNode<>root then
       begin
          node := GetFixParentNode(node);
       end;

       if endnode.ParentNode<>root then
       begin
          endnode := GetFixParentNode(endnode);
          endnode := endnode.NextSibling;
       end;
    end else
        //иначе копируем без таблицы
        begin
           if node.ParentNode<>root then
           begin
              node := GetFixParentNode(node);
              node := node.NextSibling;
           end;
           if endnode.ParentNode<>root then
           begin
              endnode := GetFixParentNode(endnode);
              endnode := endnode.PreviousSibling;
           end;
        end;

        beginnode := node;

      //Кэшируем стили
      while node<>endnode do
      begin
        nextnode := node;
        node := node.NextSibling;

        // зарегестрируем все стили которые встречаются
        CashStyles(nextnode);
     //   targetnode.ParentNode.InsertBefore(FContent.ImportNode(nextnode,true), targetnode);
      end;

    {Перенесем имена стилей и сами стили в новый документ}
    SynchronizationStyles;

    node :=beginnode;
    //Перенесем данные и переименуем их под новые стили
    while node<>endnode do
    begin
      nextnode := node;
      node := node.NextSibling;

      //Перенесем данные позднее переименовав их
   //
      importednode := FContent.ImportNode(nextnode,true);
      SynchronizationData(importednode);
      targetnode.ParentNode.InsertBefore(importednode, targetnode);
    end;

     if assigned(fpage) then
       fpage.Destroy;
     if assigned(fnextpage) then
       fnextpage.Destroy;
     CashPages();
end;

procedure TOdt.GenerateDocument(DocumentName: string = DefaultOdtFileName;
                                const DocumentPath: string = 'default');
begin
  // метод создан для задания DocumentName по-умолчанию
  // ВНИМАНИЕ!!! порядок параметров изменён!
  inherited;
end;

procedure TOdt.ShowDocument(DocumentName: string = DefaultOdtFileName;
                                     Editor: string = 'default');
begin
  // метод создан для задания DocumentName по-умолчанию
  // ВНИМАНИЕ!!! порядок параметров изменён!
  inherited;
end;

function TOdt.PrintDocument(DocumentName: string = DefaultOdtFileName): boolean;
begin
  // метод создан для задания DocumentName по-умолчанию
  Result := inherited PrintDocument(DocumentName);
end;

function TOdt.GetTable(TableName: string): TOdtTable;
begin
  Result := TOdtTable.Create(FContent,TableName);
  if not Assigned(Result) then
    MessageDlg('Ошибка','Таблица с названием "'+TableName+'" не найдена в документе',mtError,[mbOK],0);
end;

function TOdt.GetNewTableName: String;
var
    Sybols: string;
    i:integer;
begin
  i:=1;
  while TableExists('Таблица'+inttostr(i)) do            inc(i);
  Result:= 'Таблица'+inttostr(i);
end;

function TOdt.GetListofTables: TStrings;
var List: TDOMNodeList;
    i:integer;
begin
  Result := TStringList.Create;
  List:=FContent.GetElementsByTagName('table:table');
  for i:=0 to List.Count-1 do
    begin
      Result.Add(TDOMElement(List.Item[i]).AttribStrings['table:name']);
    end;
end;

function TOdt.TableExists(TableName: string): boolean;
var List: TDOMNodeList;
    i:integer;
begin
  Result:= false;
  List:=FContent.DocumentElement.GetElementsByTagName('table:table');
  for i:=0 to List.Count-1 do
    begin
      if UpperCase(TDOMElement(List.Item[i]).AttribStrings['table:name'])=UpperCase(TableName)then
        begin
          Result := true;
          break;
        end;
    end;
end;


// поиск текста во всех вложенных нодах
function TOdt.FindTextInChildNodes(Node: TDOMNode; Search: UnicodeString): TDOMNode;
var
  ChildNode: TDOMNode;
begin
  Result:= nil;
  if Node = nil then
     Node := FRoot;

  ChildNode := Node.FirstChild;
  while ChildNode<>nil do
  begin
    Result:=FindTextInChildNodes(ChildNode, Search);
    if Result<>nil then
       Exit;
    ChildNode := ChildNode.NextSibling;
  end;

  if Pos(Search, Node.TextContent)>0 then begin
//         Node.TextContent := '';
     Result:=Node;
  end;
end;




function TOdt.ContentFindFntFace(FontName: string): TDOMNode;
var FontList: TDOMNodeList;
    Attr: TDOMNamedNodeMap;
    i,j:integer;
begin
  Result:=nil;
  FontList:=FContent.GetElementsByTagName('style:font-face');
  for i:=0 to FontList.Length-1 do
    begin
      Attr:=FontList.Item[i].Attributes;
      for j:=0 to Attr.Length-1 do
        if Attr.Item[j].NodeName='style:name' then
          if UpperCase(Attr.Item[j].TextContent)=UpperCase(FontName)then
            begin
              Result:=FontList.Item[i];
              Exit;
            end;
    end;
end;

procedure TOdt.GenerateManifest;
var Root, Parent: TDOMNode;
begin
  inherited;
  // после выполнения стандартной процедуры генерации
  // метод добавляет элементы
  // специфичные для электронных таблиц
  Root := FManifest.GetElementsByTagName('manifest:manifest').Item[0];
  Parent:=FManifest.CreateElement('manifest:file-entry');
  TDOMElement(Parent).SetAttribute('manifest:version','1.2');
  TDOMElement(Parent).SetAttribute('manifest:full-path','/');
  TDOMElement(Parent).SetAttribute('manifest:media-type','application/vnd.oasis.opendocument.text');
  Root.AppendChild(Parent);
end;

procedure TOdt.GenerateContent;
var Root, Parent: TDOMNode;
begin
  // после выполнения стандартной процедуры генерации
  // метод добавляет элементы
  // специфичные для электронных таблиц
  inherited;
  Root:=FContent.DocumentElement.FindNode('office:body');
  Parent:=FContent.CreateElement('office:text');
  Root:=Root.AppendChild(Parent);

  Parent:=FContent.CreateElement('text:sequence-decls');
  Root:=Root.AppendChild(Parent);

  Parent:=FContent.CreateElement('text:sequence-decl');
  TDOMElement(Parent).SetAttribute('text:display-outline-level','0');
  TDOMElement(Parent).SetAttribute('text:name','Illustration');
  Root.AppendChild(Parent);

  Parent:=FContent.CreateElement('text:sequence-decl');
    TDOMElement(Parent).SetAttribute('text:display-outline-level','0');
    TDOMElement(Parent).SetAttribute('text:name','Table');
  Root.AppendChild(Parent);

  Parent:=FContent.CreateElement('text:sequence-decl');
    TDOMElement(Parent).SetAttribute('text:display-outline-level','0');
    TDOMElement(Parent).SetAttribute('text:name','Text');
  Root.AppendChild(Parent);

  Parent:=FContent.CreateElement('text:sequence-decl');
    TDOMElement(Parent).SetAttribute('text:display-outline-level','0');
    TDOMElement(Parent).SetAttribute('text:name','Drawing');
  {закомментировать, если необходимо чтобы в документ вставлялась 1 пустая строка}
  Root.AppendChild(Parent);
  {раскомментировать, если необходимо чтобы в документ вставлялась 1 пустая строка}
  //Root:=Root.AppendChild(Parent).ParentNode.ParentNode;
  //Parent:=FContent.CreateElement('text:p');
  //  TDOMElement(Parent).SetAttribute('text:style-name','Standard');
  //Root.AppendChild(Parent);
end;

procedure TOdt.SetDefaultTableStyles(TableName: string);
var Ts: TOdtTableStyle;
    Tw: TOdtTableColRowStyle;
    Tc: TOdtTableCellStyle;
begin
  Ts.Name:=TableName+DefaultTableStyle;
  Ts.Margin:=DefaultTableMargin;
  ts.Align:=DefaultTablePosition;
  ts.Width:=DefaultTableWidth;
  Ts.SizeCounter:=DefaultSizeCounter;

  Tw.Name:=TableName+DefaultColStyle;
  Tw.SizeCounter:=DefaultSizeCounter;
  Tw.UseOptimalColWidth:=DefaultOptimalWidth;

  Tc.Name:=TableName+DefaultCellStyle;
  Tc.border:=DefaultCellBorder;
  Tc.VerticalAlign:=DefaultVertAlign;

  SetTableStyle(Ts);
  SetTableColStyle(Tw);
  SetTableCellStyle(Tc)
end;

function TOdt.CheckDefaultTableStyles(TableName: string): boolean;
var List: TDOMNodeList;
    i:integer;
begin
Result:=false;
List:=FContent.DocumentElement.GetElementsByTagName('style:style');
for i:=0 to List.Count-1 do
  begin
    if TDOMElement(List.Item[i]).AttribStrings['style:name']=TableName+DefaultTableStyle then
      begin
         Result:=true;
         break;
      end
    else
      Result:=false;
  end;
 if Result then
   for i:=0 to List.Count-1 do
    begin
      if TDOMElement(List.Item[i]).AttribStrings['style:name']=TableName+DefaultColStyle then
        begin
           Result:=true;
           break;
        end
      else
        Result:=false;
  end;
if Result then
  for i:=0 to List.Count-1 do
    begin
      if TDOMElement(List.Item[i]).AttribStrings['style:name']=TableName+DefaultCellStyle then
        begin
           Result:=true;
           break;
        end
      else
        Result:=false;
    end;
end;

function TOdt.GetTablesCount: integer;
begin
  Result:=FContent.DocumentElement.GetElementsByTagName('table:table').Count;
end;

constructor TOdt.Create;
begin
  inherited;
end;

destructor TOdt.Destroy;
begin
  inherited;
end;

function TOdt.LoadFromFile(FileName: string): boolean;
begin
  // после выполнения стандартной процедуры загрузки
  // метод подгружает корневой элемент документа office:spreadsheet
  // специфичный для электронных таблиц
  Result := inherited LoadFromFile(FileName);
  if Result then
    FRoot := FContent.GetElementsByTagName('office:text').Item[0];
end;

function TOdt.SavetoFile(FileName: string): boolean;
begin
  GenerateDocument(ExtractFilePath(FileName),ExtractFileName(FileName));
  Result := true;
end;

function TOdt.SaveState: boolean;
begin
  if (Not Assigned(FManifest)) or (FManifest.ChildNodes.Count=0) then
    Exit;

  //удалим старые xml-файлы
  DeleteFile(TempDir+'content.xml');
  DeleteFile(TempDir+'styles.xml');
  DeleteFile(TempDir+'meta.xml');
  DeleteFile(TempDir+'settings.xml');
  DeleteFile(IncludeTrailingPathDelimiter(TempDir+'META-INF')+'manifest.xml');

  //сохраняем новые файлы
  if not DirectoryExists(TempDir+'META-INF') then CreateDir(TempDir+'META-INF');
  WriteXMLFile(FContent, TempDir+'content.xml');
  WriteXMLFile(FStyles, TempDir+'styles.xml');
  WriteXMLFile(FMeta, TempDir+'meta.xml');
  WriteXMLFile(FManifest,IncludeTrailingPathDelimiter(TempDir+'META-INF')+'manifest.xml');
  WriteXMLFile(FSettings, TempDir+'settings.xml');

  Result := true;
end;

function TOdt.SetTableStyle(TableStyle: TOdtTableStyle): boolean;
var Node, Root: TDOMNode;
begin
try
  Root:=FContent.DocumentElement.GetElementsByTagName('office:automatic-styles').Item[0];
  Node:=FContent.CreateElement('style:style');
  TDOMElement(Node).SetAttribute('style:name',TableStyle.Name);
  TDOMElement(Node).SetAttribute('style:family','table');
  Root:=Root.AppendChild(Node);
  Node:=FContent.CreateElement('style:table-properties');
  case TableStyle.SizeCounter of
    tsPercent:TDOMElement(Node).SetAttribute('style:rel-width',CurrToStr(TableStyle.Width));
    tsCM:TDOMElement(Node).SetAttribute('style:width',CurrToStr(TableStyle.Width));
  end;
  case TableStyle.Align of
    taCenter:TDOMElement(Node).SetAttribute('table:align','center');
    taLeft:TDOMElement(Node).SetAttribute('table:align','left');
    taMargins:TDOMElement(Node).SetAttribute('table:align','margins');
    taRight:TDOMElement(Node).SetAttribute('table:align','right');
  end;
TDOMElement(Node).SetAttribute('fo:margin',IntToStr(TableStyle.Margin));
Root.AppendChild(Node);
Result:=true;
except
  Result:=false
end;
end;

function TOdt.SetTableColStyle(TableColRowStyle: TOdtTableColRowStyle
  ): boolean;
var Node, Root: TDOMNode;
begin
try
  Root:=FContent.DocumentElement.GetElementsByTagName('office:automatic-styles').Item[0];
  Node:=FContent.CreateElement('style:style');
  TDOMElement(Node).SetAttribute('style:name',TableColRowStyle.Name);
  TDOMElement(Node).SetAttribute('style:family','table-column');
  Root:=Root.AppendChild(Node);
  Node:=FContent.CreateElement('style:table-column-properties');
  if TableColRowStyle.UseOptimalColWidth then
    TDOMElement(Node).SetAttribute('style:use-optimal-column-width','true')
  else
    begin
      case TableColRowStyle.SizeCounter of
        tsPercent:TDOMElement(Node).SetAttribute('style:column-width',CurrToStr(TableColRowStyle.ColWidth));
        tsCM:TDOMElement(Node).SetAttribute('style:rel-column-width',CurrToStr(TableColRowStyle.ColWidth));
      end;
    end;
  Root.AppendChild(Node);
  Result:=true;
except
  Result:=false
end;
end;

function TOdt.SetTableCellStyle(TableCellStyle: TOdtTableCellStyle): boolean;
var Root, Node: TDOMNode;
begin
try
Root:=FContent.DocumentElement.GetElementsByTagName('office:automatic-styles').Item[0];
Node:=FContent.CreateElement('style:style');
TDOMElement(Node).SetAttribute('style:name',TableCellStyle.Name);
TDOMElement(Node).SetAttribute('style:family','table-cell');
Root:=Root.AppendChild(Node);
Node:=FContent.CreateElement('style:table-cell-properties');
if Length(TableCellStyle.border)=0 then
  begin
    TDOMElement(Node).SetAttribute('fo:border-left',TableCellStyle.border_left);
    TDOMElement(Node).SetAttribute('fo:border-right',TableCellStyle.border_right);
    TDOMElement(Node).SetAttribute('fo:border-top',TableCellStyle.border_top);
    TDOMElement(Node).SetAttribute('fo:border-bottom',TableCellStyle.border_bottom);
  end
else
  TDOMElement(Node).SetAttribute('fo:border',TableCellStyle.border);
case TableCellStyle.VerticalAlign of
  vaAutomatic:TDOMElement(Node).SetAttribute('style:vertical-align','automatic');
  vaBottom:TDOMElement(Node).SetAttribute('style:vertical-align','bottom');
  vaMiddle:TDOMElement(Node).SetAttribute('style:vertical-align','middle');
  vaTop:TDOMElement(Node).SetAttribute('style:vertical-align','top');
end;
Root.AppendChild(Node);
  Result:=true;
except
  Result:=false;
end;
end;

function TOdt.GetCountPages: Integer;
begin
  Result := Length(FCashPages);
end;

function TOdt.GetPage(aIndex: integer): TPage;
var node: TDOMnode;
begin
  Result := nil;
  if aIndex=-1 then exit;
  if (aIndex <= CountPages-1)and(aIndex>=0) then
  begin
   Result:= TPage.Create;
   Result.FBreakNumber:= FCashPages[aIndex].BreakNumber;

    if Result.FBreakNumber.ParentNode <> FRoot then
    Result.FBreakNumber := GetFixParentNode(Result.FBreakNumber);

   Result.FParent := self;
   //----------------------Предыдущая страница-----------------------
   if aIndex > 0 then
   begin
     Result.FPreviousPage := FCashPages[aIndex-1].BreakNumber;
     //исправим нахождение разрыва страницы по ближе к руту если разрыв внутри таблицы
     if Result.FPreviousPage.ParentNode<>FRoot then
       Result.FPreviousPage := GetFixParentNode(Result.FPreviousPage);
     //не будем включать таблицу в состав страницы если в ней разрыв
   end
     else
       Result.FPreviousPage := GetBeginPages;
   //----------------------Следующая страница------------------------
   if aIndex < CountPages-1 then
   begin
     Result.FNextPage := FCashPages[aIndex+1].BreakNumber;
     //исправим нахождение разрыва страницы по ближе к руту если разрыв внутри таблицы
     if Result.FNextPage .ParentNode<>FRoot then
     begin
       Result.FNextPage := GetFixParentNode(Result.FNextPage);
       //Включаем таблицу в содержание данной страницы
       Result.FNextPage:=Result.FNextPage.NextSibling;
     end;
   end
     else
       Result.FNextPage := GetEndPages;
  end;
end;

procedure TOdt.RemoveStyleParagraphParametr(Style: String; Parametr: string);
var Root, Node: TDOMNode;
    List: TDOMNodeList;
    i: integer;
begin
  Root:=FContent.DocumentElement.GetElementsByTagName('office:automatic-styles').Item[0];
  List:=Root.ChildNodes;
  for i:=0 to List.Count-1 do
    if UpperCase(TDOMElement(List.Item[i]).AttribStrings['style:name'])=UpperCase(Style) then
    begin
      //Надеюсь первый в списке это именно <style:paragraph-properties
      TDOMElement(List.Item[i].ChildNodes.Item[0]).RemoveAttribute(Parametr);
    end;
end;

function TOdt.GetBeginPages: TDOMnode;
var list:TDOMnodeList;
begin
list := FContent.GetElementsByTagName('office:text').Item[0].GetChildNodes;

 if  UpperCase(TDOMElement(List.Item[0]).TagName) <> UpperCase('office:forms') then
    Result :=  List.Item[0]
    else
    if  UpperCase(TDOMElement(List.Item[1]).TagName) <> UpperCase('text:sequence-decls') then
       Result :=  List.Item[1]
       else
         Result :=  List.Item[2];
end;

function TOdt.GetEndPages: TDOMnode;
var
    list : TDOMnodeList;
    root : TDOMnode;
begin
  root := FContent.GetElementsByTagName('office:text').Item[0];
  list := root.GetChildNodes;
  Result := List[List.Count-1]
end;

function TOdt.GetFixParentNode(OrignNode: TDOMnode) : TDOMnode;
begin
  if OrignNode=nil then Exit;
  if OrignNode.ParentNode <> FRoot then
  begin
    Result := OrignNode.ParentNode ;
    Result := GetFixParentNode(Result);
  end else
        Result := OrignNode;
end;

function TOdt.ExportScripttoOO(): boolean;
var s,v: UTF8string;
    ADoc: TXMLDocument;
    Root,node:TDOMNode;
begin
  Result := false;
  //Найдем домашний каталог пользователя
  s:=GetEnvironmentVariable('HOME');
  if s='/root' then
  begin
    Showmessage('Программа запущена из под Root. Команда ConvertDOCToODT не выполнена');
    Exit;
  end;
  v:=GetOfficeVersion();
  //найдем каталог офиса
  if DirectoryExists(s+IncludeTrailingPathDelimiter('/.config/libreoffice/4/'))and
    (pos('LibreOffice',v)>0) then
  begin
    s:=s+IncludeTrailingPathDelimiter('/.config/libreoffice/4/');
  end else
    if DirectoryExists(s+IncludeTrailingPathDelimiter('/.config/openoffice/4/'))and
      (pos('OpenOffice',v)>0) then
    begin
      s:=s+IncludeTrailingPathDelimiter('/.config/openoffice/4/');
    end
      else
        s:='';

  if s='' then Exit;
  s:=s+IncludeTrailingPathDelimiter('user/basic/Standard/');
  //Cкопируем скрипт в каталог офиса(Если конечно скрипт сущществует)
  if fileExists('macrosODF.xba') then
  begin
    //Проверим в каталоге назначения есть ли скрипт
    if not fileExists(s + 'macrosODF.xba') then
      Result := CopyFile('macrosODF.xba',s + 'macrosODF.xba')
      else Result := true;
  end
   else
     Result:=false;

   if not Result then Exit;

   //Зарегестрируем\проверим есть ли такой макрос
   if FileExists(s+'script.xlb') then
   begin
      ReadXMLFile(ADoc,s+'script.xlb');
     //  <library:element library:name="macros"/>
      root:=ADoc.GetElementsByTagName('library:library').Item[0];
      node:=root.FirstChild;
      while node<>nil do
      begin
         if TDOMElement(node).AttribStrings['library:name']='macrosODF' then break;
         node:=node.NextSibling;
      end;
      if node=nil then
      begin
        node:=ADoc.CreateElement('library:element');
        TDOMElement(node).SetAttribute('library:name','macrosODF');
        root.AppendChild(node);
      end;
      WriteXMLFile(ADoc,s+'script.xlb');
      Result := true;
   end else
       Result:=false;
end;

function TOdt.InsertTable(Cols, Rows: integer; TableName : string) : TOdtTable;
begin
  Result:=InsertTable(FRoot.LastChild,Cols, Rows, TableName);
end;

function TOdt.InsertTable(InsertinBefore: TDOMnode; Cols, Rows: integer
  ) : TOdtTable;
var
  tablename : string;
begin
  tablename := SearchFreeStyleName(DefaultTableName+'1');
  Result:=InsertTable(InsertinBefore, Cols, Rows, tablename);
end;

function TOdt.InsertTable(Cols, Rows: integer) : TOdtTable;
var
  tablename : string;
begin
  tablename := SearchFreeStyleName(DefaultTableName+'1');
  Result:=InsertTable(FRoot, Cols, Rows, tablename);
end;

function TOdt.InsertTable(InsertinBefore: TDOMnode; Cols, Rows: integer;
  TableName: string) : TOdtTable;
var Root, Node: TDOMNode;
    i,j:integer;
    s: UnicodeString;
begin

  //добавление пустой таблицы  в документ
  if not CheckDefaultTableStyles(TableName) then
     SetDefaultTableStyles(TableName);

  // Вставлять в элемент верхнего уровня
  while (InsertinBefore.ParentNode<>nil) do begin
        if InsertinBefore.ParentNode.NodeName ='office:text' then break;
        InsertinBefore:=InsertinBefore.ParentNode;
  end;

      Root:=FRoot;
      Node:=FContent.CreateElement('table:table');
      TDOMElement(Node).SetAttribute('table:name',TableName);
      TDOMElement(Node).SetAttribute('table:style-name',TableName+DefaultTableStyle);
      Root:=InsertinBefore.ParentNode.InsertBefore(Node,InsertinBefore);
      Node:=FContent.CreateElement('table:table-column');
      TDOMElement(Node).SetAttribute('table:style-name',TableName+DefaultColStyle);
      TDOMElement(Node).SetAttribute('table:number-columns-repeated',IntToStr(Cols));
      Root:=Root.AppendChild(Node);//Root = table:table-column
      //вставляем строки и ячейки
      for i:=1 to Rows do
        begin
          Root:=Root.ParentNode;//Root = table:table
          Node:=FContent.CreateElement('table:table-row');
          Root:=Root.AppendChild(Node); //Root = table:table-row
          for j:=1 to Cols do
            begin
              Node:=FContent.CreateElement('table:table-cell');
              TDOMElement(Node).SetAttribute('table:style-name',TableName+DefaultCellStyle);
              TDOMElement(Node).SetAttribute('office:value-type','string');
              Root:=Root.AppendChild(Node);//Root = table:table-cell
              Node:=FContent.CreateElement('text:p');
              TDOMElement(Node).SetAttribute('text:style-name','Standard');
              Root:=Root.AppendChild(Node).ParentNode.ParentNode;//Root = table:table-row
            end;
        end;
(*
      //добавление пустой таблицы  в документ
      if CheckDefaultTableStyles(TableName) then
      begin
          Root:=FRoot;
          Node:=FContent.CreateElement('table:table');
          TDOMElement(Node).SetAttribute('table:name',TableName);
          TDOMElement(Node).SetAttribute('table:style-name',TableName+DefaultTableStyle);
          Root:=InsertinBefore.ParentNode.InsertBefore(Node,InsertinBefore);
          Node:=FContent.CreateElement('table:table-column');
          TDOMElement(Node).SetAttribute('table:style-name',TableName+DefaultColStyle);
          TDOMElement(Node).SetAttribute('table:number-columns-repeated',IntToStr(Cols));
          Root:=Root.AppendChild(Node);//Root = table:table-column
          //вставляем строки и ячейки
          for i:=1 to Rows do
            begin
              Root:=Root.ParentNode;//Root = table:table
              Node:=FContent.CreateElement('table:table-row');
              Root:=Root.AppendChild(Node); //Root = table:table-row
              for j:=1 to Cols do
                begin
                  Node:=FContent.CreateElement('table:table-cell');
                  TDOMElement(Node).SetAttribute('table:style-name',TableName+DefaultCellStyle);
                  TDOMElement(Node).SetAttribute('office:value-type','string');
                  Root:=Root.AppendChild(Node);//Root = table:table-cell
                  Node:=FContent.CreateElement('text:p');
                  TDOMElement(Node).SetAttribute('text:style-name','Standard');
                  Root:=Root.AppendChild(Node).ParentNode.ParentNode;//Root = table:table-row
                end;
            end;
      end else
        begin
          SetDefaultTableStyles(TableName);
          InsertTable(Cols, Rows, TableName);
        end;
*)
  Result  := GetTable(TableName);

end;

procedure TOdt.DeleteTable(TableName: string);
var List: TDOMNodeList;
    i:integer;
    node: TDOMNode;
begin
  List:=FContent.DocumentElement.GetElementsByTagName('table:table');
  for i:=0 to List.Count-1 do
    begin
      if UpperCase(TDOMElement(List.Item[i]).AttribStrings['table:name'])=UpperCase(TableName)then
        begin
          node := List.Item[i];
          node.parentNode.RemoveChild(node);
          break;
        end;
    end;

end;

procedure TOdt.DeleteTable(aTable: TOdtTable);
begin
   if assigned(aTable) then
   with aTable do
   begin
      FRoot.parentNode.RemoveChild(FRoot);
     Destroy;
   end;

end;

function TOdt.ShowPartOfDocument(Doc: TFileType  = ftContent;editor: string = 'default'): boolean;
var Proc: TProcess;
    f_name: string;
begin
  Result:=False;
  if DocumentLoaded then
  begin
    SaveState();

    case Doc of
      ftStyles: f_name:= TempDir+'styles.xml';
      ftContent: f_name:= TempDir+'content.xml' ;
      ftManifest: f_name:= IncludeTrailingPathDelimiter(TempDir+'META-INF')+'manifest.xml';
      ftMeta: f_name:= TempDir+'meta.xml';
      ftSettings: f_name:= TempDir+'settings.xml';
    end;

     if editor='default' then begin
      if not OpenDocument(f_name) then
        if not OpenURL(f_name) then
          begin
            MessageDlg('Ошибка','Не удалось открыть файл "'+f_name+'"',mtError,[mbOK],0);
           sleep(2000);
          end;

     end
     else begin
       try
         Proc:=TProcess.Create(nil);
         Proc.Executable:=editor;
         Proc.Parameters.Append(TempDir+f_name);
         Proc.Options:=[poWaitOnExit];
         Proc.ShowWindow:=swoShowMaximized;
         Proc.Execute;
       finally
         if Proc.WaitOnExit then Proc.Free;
       end;
     end;
  end;
end;

{ TOdtTable }

// проверка изменений в полях типа boolean
function Changed(B,OB:boolean):boolean; overload;
begin
  if OB=Boolean(-1) then
    if B<>Boolean(-1) then exit(true)
    else exit(false)
  else
    if B=Boolean(-1) then exit(false)
    else if B<>OB then exit(true);
end;

// получить индекс записи массива, входящей в s
// ! требуется полное совпадение
function GetIdxF(s:string; arr: array of string):integer;
var i:integer;
begin
  for i:=0 to high(arr) do
    if arr[i]=s then exit(i);
end;

// получить индекс записи массива, входящей в s
// ! только для комбинированных записей вида '10cm' или '15%'
function GetIdx(s:string; arr: array of string):integer;
var i:integer;
begin
  for i:=0 to high(arr) do
    if UTF8Pos(arr[i],s)>0 then exit(i);
end;

// получение значения цифровой части из смешанной строки, откидывая суффиксы
function GetValue(s: string):double;
var valid_sym  : string = '0123456789.-';
    j,r:integer;
    d: string;
begin
  for j:=1 to UTF8Length(s) do
    if UTF8Pos(s[j],valid_sym)=0 then begin
      r:=j;
      break;
    end;
  d:=UTF8Copy(s,1,r-1);
  Result:=LazUtilities.StrToDouble(d);
end;

// закодирование типа TLength в строку
function EncodeSizes(v:TLength):string;
var s:string;
begin
  s:=FloatToStr(v.Value)+MeasureArr[ord(v.Measure)];
  Result:=s;
end;

// раскодирование типа TLength из строки
function DecodeSizes(s:string):TLength;
begin
  Result.Value:=GetValue(s);
  Result.Measure:=TMeasure(GetIdx(s,MeasureArr));
end;

// получим символ из имени стиля КОЛОНКИ
function TOdtTable.GetSymbol(s:string):string;
var pos:integer;
begin
  pos:=UTF8Pos('.',s);
  if (pos>0) and (UTF8Copy(s,1,pos-1)=Name) then
    Result:=UTF8Copy(s,pos+1,UTF8Length(s)-pos)
  else exit('');
end;

// получить свойства таблицы
procedure TOdtTable.GetTableProperties(var P: TOdtTableProperties);
var AutoStylesNode, PropNode: TDOMNode;
    List: TDOMNodeList;
    i, c: integer;
    v, nn: string;

begin
  AutoStylesNode:=TXMLDocument(Document).FirstChild.FindNode('office:automatic-styles');
  if assigned(AutoStylesNode) and AutoStylesNode.HasChildNodes then
    List:=AutoStylesNode.ChildNodes
  else exit;

  for i:=0 to List.Count-1 do
      if (List.Item[i].NodeName='style:style') and
        (TDOMElement(List.Item[i]).GetAttribute('style:name')=Name) and
        (TDOMElement(List.Item[i]).GetAttribute('style:family')='table') then
        begin
          PropNode:=List.Item[i].FindNode('style:table-properties');
          if assigned(PropNode) and PropNode.HasAttributes then
            for c:=0 to PropNode.Attributes.Length-1 do begin
              v:=PropNode.Attributes[c].NodeValue;
              nn:=PropNode.Attributes[c].NodeName;
              if nn='style:width' then P.Width:=DecodeSizes(v);
              if nn='fo:background-color' then P.BackgroundColor:=HtmlHexToColor(v);
              if nn='fo:break-after' then P.BreakAfter:=TBreak(GetIdxF(v,BreakArr));
              if nn='fo:break-before' then P.BreakBefore:=TBreak(GetIdxF(v,BreakArr));
              if nn='fo:keep-with-next' then P.KeepWithNext:=TKeepWithNext(GetIdxF(v,KeepWithNextArr));
              //if nn='fo:margin' then P.Margin:=DecodeSizes(v);
              if nn='fo:margin-bottom' then P.MarginBottom:=TNonNegativeLength(DecodeSizes(v));
              if nn='fo:margin-left' then P.MarginLeft:=DecodeSizes(v);
              if nn='fo:margin-right' then P.MarginRight:=DecodeSizes(v);
              if nn='fo:margin-top' then P.MarginTop:=TNonNegativeLength(DecodeSizes(v));
              if nn='style:may-break-between-rows' then P.MayBreakBetweenRows:=StrToBool(v);
              if nn='style:page-number' then P.PageNumber:=StrToPageNumber(v);
              if nn='style:rel-width' then P.RelWidth:=DecodeSizes(v);
              if nn='style:shadow' then P.Shadow:=v;
              //if nn='style:writing-mode' then P.WritingMode:=TWritingMode(GetIdxF(v,WritingModeArr));
              if nn='table:align' then P.Align:=TTableAlign(GetIdxF(v,TableAlignArr));
              if nn='table:border-model' then P.BorderModel:=TBorderModel(GetIdxF(v,BorderModelArr));
              //if nn='table:display' then P.Display:=StrToBool(v);

            end;

          //The <style:table-properties> element has the following child element:
          //<style:background-image> 17.3.
        end;
end;

procedure TOdtTable.GetColumns;
var List: TDOMNodeList;
    i, c: integer;
    v, nn: string; // v = value; nn = nodeName
    cc: integer; // cc = clone count
    d: integer; // d = donor
    nc: integer; // nc = new column

begin
  FColumns.Clear;
  List:=FRoot.ChildNodes;
  for i:=0 to List.Count-1 do
    if List.Item[i].NodeName='table:table-column' then begin
      nc:=FColumns.Add(TOdtTableColumn.Create);
      cc:=0;
      for c:=0 to List.Item[i].Attributes.Length-1 do begin
        v:=List.Item[i].Attributes[c].NodeValue;
        nn:=List.Item[i].Attributes[c].NodeName;
        if nn='table:default-cell-style-name' then FColumns[nc].DefaultCellStyleName:=v;
        if nn='table:number-columns-repeated' then begin
          cc:=StrToInt(v);
          d:=nc;
        end;
        if nn='table:style-name' then FColumns[nc].StyleName:=v;
        if nn='table:visibility' then FColumns[nc].Visibility:=
                               TTableVisibility(GetIdxF(v,TableVisibilityArr));
        //if nn='xml:id' then P.XMLId:=v;
      end;
      while cc>1 do begin
        nc:=FColumns.Add(TOdtTableColumn.Create);
        FColumns[nc]:=FColumns[d];
        dec(cc);
      end;
    end;
end;

// получить строки таблицы
procedure TOdtTable.GetRows;
var List: TDOMNodeList;
    i, c: integer;
    v, nn: string; // v = value; nn = nodeName
    cc: integer; // cc = clone count
    d: integer; // d = donor
    nr: integer; // nr = new row

begin
  FRows.Clear;
  List:=FRoot.ChildNodes;
  for i:=0 to List.Count-1 do
    if List.Item[i].NodeName='table:table-row' then begin
      nr:=FRows.Add(TOdtTableRow.Create);
      cc:=0;
      for c:=0 to List.Item[i].Attributes.Length-1 do begin
        v:=List.Item[i].Attributes[c].NodeValue;
        nn:=List.Item[i].Attributes[c].NodeName;
        if nn='table:default-cell-style-name' then FRows[nr].DefaultCellStyleName:=v;
        if nn='table:number-rows-repeated' then begin
          cc:=StrToInt(v);
          d:=nr;
        end;
        if nn='table:style-name' then FRows[nr].StyleName:=v;
        if nn='table:visibility' then FRows[nr].Visibility:=
                               TTableVisibility(GetIdxF(v,TableVisibilityArr));

        //if nn='xml:id' then P.XMLId:=v;
      end;
      while cc>1 do begin     // каждой строке создаём объект, даже если они одинаковы
        nr:=FRows.Add(TOdtTableRow.Create);
        FRows[nr]:=FRows[d];
        dec(cc);
      end;
    end;
end;

// получить свойства колонок таблицы
procedure TOdtTable.GetColsProperties;
var AutoStylesNode, PropNode: TDOMNode;
    List: TDOMNodeList;
    i, c: integer;
    v, nn: string;  // v = value; nn = nodeName
    nc: integer; // nc = new column

  procedure UnPackColsProperties;
  var i:integer;
  begin
    for i:=0 to FColsProperties.Count-1 do
    begin
      if TOdtTableColumnProperties(FColsProperties[i]).Symbol<>NumToSym(i+1) then
        if i>0 then begin
          FColsProperties.Insert(i,TOdtTableColumnProperties.Create);
          FColsProperties[i].Width       := FColsProperties[i-1].Width;
          //FColsProperties[i].BreakAfter  := FColsProperties[i-1].BreakAfter;
          //FColsProperties[i].BreakBefore := FColsProperties[i-1].BreakBefore;
          //FColsProperties[i].RelWidth    := FColsProperties[i-1].RelWidth;
          //FColsProperties[i].UseOptimalSize := FColsProperties[i-1].UseOptimalSize;
          FColsProperties[i].Symbol      := NumToSym(i+1);
        end;
    end;

  end;

begin
  AutoStylesNode:=TXMLDocument(Document).FirstChild.FindNode('office:automatic-styles');
  if assigned(AutoStylesNode) and AutoStylesNode.HasChildNodes then
    List:=AutoStylesNode.ChildNodes
  else exit;

  FColsProperties.Clear;
  for i:=0 to List.Count-1 do
  begin
      if (List.Item[i].NodeName='style:style') and
        (TDOMElement(List.Item[i]).GetAttribute('style:family')='table-column')
        and InTable(TDOMElement(List.Item[i]).GetAttribute('style:name')) then
        begin
          PropNode:=List.Item[i].FindNode('style:table-column-properties');
          if assigned(PropNode) and PropNode.HasAttributes then begin
            nc:=FColsProperties.Add(TOdtTableColumnProperties.Create);
            FColsProperties[nc].Symbol:=
              GetSymbol(TDOMElement(List.Item[i]).GetAttribute('style:name'));
            for c:=0 to PropNode.Attributes.Length-1 do begin
              v:=PropNode.Attributes[c].NodeValue;
              nn:=PropNode.Attributes[c].NodeName;
              if nn='style:column-width' then
                FColsProperties[nc].Width:=DecodeSizes(v);
              //if nn='fo:break-after' then
                //FColsProperties[nc].BreakAfter:=TBreak(GetIdxF(v,BreakArr));
              //if nn='fo:break-before' then
                //FColsProperties[nc].BreakBefore:=TBreak(GetIdxF(v,BreakArr));
              //if nn='style:rel-column-width' then
                //FColsProperties[nc].RelWidth:=DecodeSizes(v);
              //if nn='style:use-optimal-column-width' then
                //FColsProperties[nc].UseOptimalSize:=TExtBool(GetIdxF(v,ExtBoolArr));
            end;
          end;
        end;
  end;
  UnPackColsProperties;
end;

function TOdtTable.GetColCount: integer;
var List: TDOMNodeList;
    i,r:integer;
begin
  r:=0;
  List:=FRoot.ChildNodes;
  for i:=0 to List.Count-1 do
    if List.Item[i].NodeName='table:table-column' then
      if TDOMElement(List.Item[i]).hasAttribute('table:number-columns-repeated') then
        r:=r+StrToInt(TDOMElement(List.Item[i]).AttribStrings['table:number-columns-repeated'])
      else
        inc(r);
  Result:=r;
end;

function TOdtTable.GetRowCount: integer;
var List: TDOMNodeList;
    i:integer;
begin
  Result:=0;
  List:=FRoot.ChildNodes;
  for i:=0 to List.Count-1 do
    if List.Item[i].NodeName='table:table-row' then inc(Result);
end;

// установка свойств таблицы
procedure TOdtTable.SetTableProperties;
var AutoStylesNode, PropNode: TDOMNode;
    List: TDOMNodeList;
    i:integer;
    P: TOdtTableProperties;

  procedure SetAt(N,V:string);
  begin TDOMElement(PropNode).SetAttribute(N,V); end;

begin
  AutoStylesNode:=TXMLDocument(Document).FirstChild.FindNode('office:automatic-styles');
  if assigned(AutoStylesNode) and AutoStylesNode.HasChildNodes then
    List:=AutoStylesNode.ChildNodes
  else exit;

  for i:=0 to List.Count-1 do
      if (List.Item[i].NodeName='style:style') and
        (TDOMElement(List.Item[i]).GetAttribute('style:name')=Name) and
        (TDOMElement(List.Item[i]).GetAttribute('style:family')='table') then
        begin
          PropNode:=List.Item[i].FindNode('style:table-properties');
          if assigned(PropNode) then begin
            P:=TOdtTableProperties.Create;
            GetTableProperties(P);
            if Changed(FProperties.Width,P.Width) then
              SetAt('style:width', EncodeSizes(FProperties.FWidth));
            if (FProperties.BackgroundColor<>P.BackgroundColor) then
              SetAt('fo:background-color', ColorToHtmlHex(FProperties.BackgroundColor));
            if FProperties.BreakAfter<>P.BreakAfter then
              SetAt('fo:break-after', BreakArr[ord(FProperties.BreakAfter)]);
            if FProperties.BreakBefore<>P.BreakBefore then
              SetAt('fo:break-before', BreakArr[ord(FProperties.BreakBefore)]);
            if FProperties.KeepWithNext<>P.KeepWithNext then
              SetAt('fo:keep-with-next', KeepWithNextArr[ord(FProperties.KeepWithNext)]);
            //if Changed(FProperties.Margin,P.Margin) then
              //SetAt('fo:margin', EncodeSizes(FProperties.Margin));
            if Changed(FProperties.MarginBottom,P.MarginBottom) then
              SetAt('fo:margin-bottom', EncodeSizes(FProperties.MarginBottom));
            if Changed(FProperties.MarginLeft,P.MarginLeft) then
              SetAt('fo:margin-left', EncodeSizes(FProperties.MarginLeft));
            if Changed(FProperties.MarginRight,P.MarginRight) then
              SetAt('fo:margin-right', EncodeSizes(FProperties.MarginRight));
            if Changed(FProperties.MarginTop,P.MarginTop) then
              SetAt('fo:margin-top', EncodeSizes(FProperties.MarginTop));
            if Changed(FProperties.MayBreakBetweenRows,P.MayBreakBetweenRows) then
              SetAt('style:may-break-between-rows', BoolToStr(FProperties.MayBreakBetweenRows,'true','false'));
            if Changed(FProperties.PageNumber,P.PageNumber) then
              SetAt('style:page-number', PageNumberToStr(FProperties.PageNumber));
            if Changed(FProperties.RelWidth,P.RelWidth) then
              SetAt('style:rel-width', FloatToStr(FProperties.RelWidth.Value));
            if FProperties.Shadow<>P.Shadow then
              SetAt('style:shadow', FProperties.Shadow);
            //if FProperties.WritingMode<>P.WritingMode then
              //SetAt('style:writing-mode', WritingModeArr[ord(FProperties.WritingMode)]);
            if FProperties.Align<>P.Align then
              SetAt('table:align', TableAlignArr[ord(FProperties.Align)]);
            if FProperties.BorderModel<>P.BorderModel then
              SetAt('table:border-model', BorderModelArr[ord(FProperties.BorderModel)]);
            //if Changed(FProperties.Display,P.Display) then
              //SetAt('table:display', BoolToStr(FProperties.Display,'true','false'));
            FreeAndNil(P);
          end;
        end;
end;

// установка свойств колонок таблицы
procedure TOdtTable.SetColsProperties;
var AutoStylesNode, n, x: TDOMNode;
    i,c:integer;
    ColProps: array of string;

  function AddNode(Symbol:string):TDOMNode;
  var Node,PrNode:TDOMNode;
      j: integer;
  begin
    j:=StrToInt(SymToNum(Symbol));
    Node:=TXMLDocument(Document).CreateElement('style:style');
    TDOMElement(Node).SetAttribute('style:name',Name+'.'+Symbol);
    TDOMElement(Node).SetAttribute('style:family','table-column');
    PrNode:=TXMLDocument(Document).CreateElement('style:table-column-properties');
    //if FColsProperties.Items[i].UseOptimalSize=ebTrue then
      //TDOMElement(PrNode).SetAttribute('style:use-optimal-column-width',ExtBoolArr[ord(ebTrue)])
    //else
    TDOMElement(PrNode).SetAttribute('style:column-width',EncodeSizes(FColsProperties.Items[j-1].Width));
    //TDOMElement(PrNode).SetAttribute('style:rel-column-width',EncodeSizes(FColsProperties.Items[j-1].RelWidth));
    //if FColsProperties.Items[i].BreakAfter<>bNil then
    //  TDOMElement(PrNode).SetAttribute('fo:break-after',
    //    BreakArr[ord(FColsProperties.Items[i].BreakAfter)]);
    //if FColsProperties.Items[i].BreakBefore<>bNil then
    //  TDOMElement(PrNode).SetAttribute('fo:break-before',
    //    BreakArr[ord(FColsProperties.Items[i].BreakBefore)]);

    Node.AppendChild(PrNode);

    if n.NextSibling<>nil then
      n:=AutoStylesNode.InsertBefore(Node,n.NextSibling)
    else
      n:=AutoStylesNode.AppendChild(Node);

    Result:=Node;
  end;

begin
  AutoStylesNode:=TXMLDocument(Document).FirstChild.FindNode('office:automatic-styles');
  if not assigned(AutoStylesNode) then exit;

  // сначла удалим все стили колонок таблицы
  n:=AutoStylesNode.LastChild;
  while n<>nil do begin
    x:=n.PreviousSibling;
    if (n.NodeName='style:style') and
      (TDOMElement(n).GetAttribute('style:family')='table-column') and
      InTable(TDOMElement(n).GetAttribute('style:name')) then n.Destroy;
    n:=x;
  end;

  // найдём ноду стиля таблицы (вставлять будем после неё)
  n:=AutoStylesNode.FirstChild;
  while n<>nil do begin
    if (n.NodeName='style:style') and
      (TDOMElement(n).GetAttribute('style:family')='table') and
      (TDOMElement(n).GetAttribute('style:name')=Name) then break;
    n:=n.NextSibling;
  end;

  SetLength(ColProps,0);
  for i:=FColsProperties.Count-1 downto 0 do
    if i=0 then begin
      SetLength(ColProps,Length(ColProps)+1);
      ColProps[high(ColProps)]:=FColsProperties.Items[i].FSymbol;
    end
    else begin
      if FColsProperties.Items[i].Equal(FColsProperties.Items[i-1]) then begin
        // найти все колонки с этим стилем и изменить ссылку на предыдущий стиль
        for c:=0 to FColumns.Count-1 do
          if (FColumns.Items[c].StyleName=Name+'.'+NumToSym(i+1)) then
            FColumns.Items[c].StyleName:=Name+'.'+NumToSym(i);
        continue
      end
      else begin
        SetLength(ColProps,Length(ColProps)+1);
        ColProps[high(ColProps)]:=FColsProperties.Items[i].FSymbol;
      end;
    end;

  // теперь имея список нужных для записи нод - запишем их в FContent
  for i:=length(ColProps) downto 1 do AddNode(ColProps[i-1]);

end;

procedure TOdtTable.SetColumns;
var List: TDOMNodeList;
    i:integer;
    Node:TDOMNode;

begin
// 1. удалить все старые записи о колонках
// 2. записать в хмл данные из FColumns

  // удаляем
  List:=FRoot.ChildNodes;
  for i:=List.Count-1 downto 0 do
    if List.Item[i].NodeName='table:table-column' then
      List.Item[i].Destroy;

  // записываем
  for i:=FColumns.Count-1 downto 0 do begin
    Node:=TXMLDocument(Document).CreateElement('table:table-column');
    TDOMElement(Node).SetAttribute('table:style-name',FColumns.Items[i].StyleName);
    FRoot.InsertBefore(Node,FRoot.FirstChild);
  end;
end;

//procedure TOdtTable.GetColumns(var P: TOdtTableColumns);
//var List: TDOMNodeList;
//    i, c: integer;
//    v, nn: string;
//    CloneCount: integer;
//    Donor: integer;
//begin
//  SetLength(P,0);
//  List:=FRoot.ChildNodes;
//  for i:=0 to List.Count-1 do
//    if List.Item[i].NodeName='table:table-column' then begin
//      SetLength(P,Length(P)+1);
//      P[high(P)]:=TOdtTableColumn.Create;
//      CloneCount:=0;
//      for c:=0 to List.Item[i].Attributes.Length-1 do begin
//        v:=List.Item[i].Attributes[c].NodeValue;
//        nn:=List.Item[i].Attributes[c].NodeName;
//        if nn='table:default-cell-style-name' then P[high(P)].DefaultCellStyleName:=v;
//        if nn='table:number-columns-repeated' then begin
//          CloneCount:=StrToInt(v);
//          Donor:=high(P);
//        end;
//        if nn='table:style-name' then P[high(P)].StyleName:=v;
//        if nn='table:visibility' then P[high(P)].Visibility:=TTableVisibility(GetIdxF(v,TableVisibilityArr));
//        //if nn='xml:id' then P.XMLId:=v;
//      end;
//      while CloneCount>1 do begin
//        SetLength(P,Length(P)+1);
//        P[high(P)]:=TOdtTableColumn.Create;
//        P[high(P)]:=P[Donor];
//        dec(CloneCount);
//      end;
//    end;
//end;

procedure TOdtTable.SetColCount(AColCount: integer);
var i:integer;
begin
 if (AColCount-ColCount)>0 then //наращиваем количество столбцов
   begin
     for i:=1 to (AColCount-ColCount) do
       AppendColumn('');
   end
 else
   if (AColCount-ColCount)<0 then //удаляем столбцы с конца
     for i:=1 to abs((AColCount-ColCount)) do
       RemoveColumn;
end;

procedure TOdtTable.SetRowCount(ARowCount: integer);
var i:integer;
begin
 if (ARowCount-RowCount)>0 then //наращиваем количество строк
   begin
     for i:=1 to (ARowCount-RowCount) do
       InsertRow;//AppendRow('');
   end
 else
   if (ARowCount-RowCount)<0 then //удаляем строки с конца
     for i:=1 to abs((ARowCount-RowCount)) do
       RemoveRow;
end;

function TOdtTable.GetCells(ACol, ARow: Integer): string;
var node: TDOMnode;
begin
 if (ACol<0) then Exit;
 if (ARow<0) then Exit;
  node:= GetCellNode(ACol, ARow);
  if not Assigned(node) then Exit;
  //Если ячейка состоит из кучи слов со стилями,
  //собираем все слова в предложение
  if node.ParentNode.ChildNodes.Count>0 then
    Result := Assembledtext(node.ParentNode)
    else
      Result:=node.TextContent;

end;

procedure TOdtTable.SetCells(ACol, ARow: Integer; const AValue: string);
var node,nextnode,rnode: TDOMnode;
    i: integer;
begin
  if (ACol<0) then Exit;
  if (ARow<0) then Exit;
  node:= GetCellNode(ACol, ARow);
  if not Assigned(node) then Exit;
  //обычный текст
  if node.NodeName='text:p' then
  begin
    nextnode:= node.FirstChild;
    //Перед присвоением проверяем а нету ли других текстов в ячейке
    while nextnode<>nil do
    begin
       rnode:= nextnode;
       nextnode:= rnode.NextSibling;
       rnode.ParentNode.RemoveChild(rnode);

    end;
   { if node.ParentNode.ChildNodes.Count>0 then
      for i:=0 to node.ParentNode.ChildNodes.Count-1 do
          if node.ParentNode.ChildNodes[i]<>node then
               node.ParentNode.RemoveChild(node.ParentNode.ChildNodes[i]); }
    node.TextContent := AValue;

  end;
  //В ячейке находится текст с нумерацией
  if node.NodeName='text:list' then
  begin

  end;
end;
 // <table:table-header-rows>

 function TOdtTable.HeaderRowsExists(): Boolean;
 var i:integer;
     List: TDOMNodeList;
     Node: TDOMNode;
 begin
   Result:=false;
   List:=FRoot.ChildNodes;
   for i:=0 to List.Count-1 do
     begin
       if List.Item[i].NodeName='table:table-header-rows' then
         begin
           Result:=true;
           Exit;
         end;
     end;
 end;

function TOdtTable.GetHeaderRowsNode(ACol: Integer): TDOMNode;
var i,j:integer;
    List: TDOMNodeList;
    Node: TDOMNode;
begin
  List:=FRoot.ChildNodes;
  for i:=0 to List.Count-1 do
    begin
      if List.Item[i].NodeName='table:table-header-rows' then
        begin
          Node:=List.Item[i];
          for j:=0 to Node.ChildNodes.Count-1 do
            begin
              if Node.ChildNodes.Item[j].NodeName='table:table-row' then
                begin
                      Node:=Node.ChildNodes.Item[j];
                      break;
                end;
            end;
        end;
    end;
  if Node=nil then exit;
 Result:=Node.ChildNodes.Item[ACol].FirstChild;
end;

function TOdtTable.GetHeaderRows(ACol: Integer): string;
var node: TDOMnode;
begin
  if not HeaderRowsExists then Exit;
  node:= GetHeaderRowsNode(ACol);
  if not Assigned(node) then Exit;
  //Если ячейка состоит из кучи слов со стилями,
  //собираем все слова в предложение
  if node.ParentNode.ChildNodes.Count>0 then
    Result := Assembledtext(node.ParentNode)
    else
      Result:=node.TextContent;

end;

procedure TOdtTable.SetHeaderRows(ACol: Integer; const AValue: string);
var node,nextnode,rnode: TDOMnode;
    i: integer;
begin
  if not HeaderRowsExists then Exit;
  node:= GetHeaderRowsNode(ACol);
  if not Assigned(node) then Exit;
  //обычный текст
  if node.NodeName='text:p' then
  begin
    nextnode:= node.FirstChild;
    //Перед присвоением проверяем а нету ли других текстов в ячейке
    while nextnode<>nil do
    begin
       rnode:= nextnode;
       nextnode:= rnode.NextSibling;
       rnode.ParentNode.RemoveChild(rnode);

    end;
   { if node.ParentNode.ChildNodes.Count>0 then
      for i:=0 to node.ParentNode.ChildNodes.Count-1 do
          if node.ParentNode.ChildNodes[i]<>node then
               node.ParentNode.RemoveChild(node.ParentNode.ChildNodes[i]); }
    node.TextContent := AValue;

  end;
  //В ячейке находится текст с нумерацией
  if node.NodeName='text:list' then
  begin

  end;
end;

function TOdtTable.GetCellNode(ACol, ARow: Integer): TDOMNode;
var i,j:integer;
    List: TDOMNodeList;
    Node: TDOMNode;
begin
  List:=FRoot.ChildNodes;
  j:=-1;
  for i:=0 to List.Count-1 do
    begin
      if List.Item[i].NodeName='table:table-row' then
        begin
          inc(j);
          if j=ARow then
            begin
              Node:=List.Item[i];//нашли узел необходимой строки
              break;
            end;
        end;
    end;
if Node=nil then Exit;
Result:=Node.ChildNodes.Item[ACol].FirstChild;
end;

// вспомогательная функция - принимает имя стиля и отвечает принадлежит ли он
// к текущей таблице
function TOdtTable.InTable(s:string):boolean;
var pos:integer;
begin
  {Result:=false;
  pos:=UTF8Pos('.',s);
  if (pos>0) and (UTF8Copy(s,1,pos-1)=Name) then Result:=true;   }
  //Сделал проще, вообще такое имя в строке встречается?
  Result:=(UTF8Pos(Name,s)>0);
end;

constructor TOdtTable.Create(XMLDoc:TXMLDocument; TableName:string);
begin
  inherited Create(XMLDoc,TableName);
  FDefTextStyle:=DefaultTextStyle;
  FreeAndNil(FProperties);
  FProperties:=TOdtTableProperties.Create;
  FreeAndNil(FColumns);
  FColumns:=TOdtTableColumns.Create;
  FreeAndNil(FColsProperties);
  FColsProperties:=TOdtTableColsProperties.Create;
  // получим свойства таблицы (для доступа и редактирования)
  GetTableProperties(FProperties);
  // получим колонки
  GetColumns;
  // получим свойства колонок
  GetColsProperties;
end;

destructor TOdtTable.Destroy;
begin
  FProperties:=nil;
  FColumns:=nil;
  FColsProperties:=nil;
  inherited;
end;

procedure TOdtTable.SetTextStyle(StyleName, FontName: string; FontSize:integer;
  FontStyles: TFontStyles; TextPosition: TTextPosition);
var Root, Node: TDomNode;
begin
  Root:=FDocument.DocumentElement.FindNode('office:automatic-styles');
  Node:=FDocument.CreateElement('style:style');
    TDOMElement(Node).SetAttribute('style:name',StyleName);
    TDOMElement(Node).SetAttribute('style:family','paragraph');
    TDOMElement(Node).SetAttribute('style:parent-style-name','Standard');
  Root:=Root.AppendChild(Node);
  Node:=FDocument.CreateElement('style:paragraph-properties');
  case TextPosition of
    tpCenter:TDOMElement(Node).SetAttribute('fo:text-align','center');
    tpJustify:TDOMElement(Node).SetAttribute('fo:text-align','justify');
    tpLeft:TDOMElement(Node).SetAttribute('fo:text-align','start');
    tpRight:TDOMElement(Node).SetAttribute('fo:text-align','end');
  end;
  TDOMElement(Node).SetAttribute('style:justify-single-word','false');
  Root.AppendChild(Node);
  Node:=FDocument.CreateElement('style:text-properties');
  if not CheckFontName(FontName) then InsertNewFont(FontName);
  TDOMElement(Node).SetAttribute('style:font-name',FontName);
  TDOMElement(Node).SetAttribute('fo:font-size',IntToStr(FontSize));
  if ftItalic in FontStyles then
   begin
     TDOMElement(Node).SetAttribute('fo:font-style','italic');
     TDOMElement(Node).SetAttribute('style:font-style-complex','italic');
   end;
 if ftBold in FontStyles then
   begin
     TDOMElement(Node).SetAttribute('fo:font-weight','bold');
     TDOMElement(Node).SetAttribute('style:font-weight-complex','bold');
   end;
 if ftUnderline in FontStyles then
   begin
     TDOMElement(Node).SetAttribute('style:text-underline-style','solid');
     TDOMElement(Node).SetAttribute('style:text-underline-width','auto');
     TDOMElement(Node).SetAttribute('style:text-underline-color','font-color');
   end;
 Root.AppendChild(Node);
end;

procedure TOdtTable.ApplyTextStyle(ACol, ARow: Integer; AStyle: string);
begin
  TDOMElement(GetCellNode(ACol,ARow)).AttribStrings['text:style-name']:=AStyle;
end;

procedure TOdtTable.AppendColumn(TextStyle:string);
var Root, Node, FirstRowNode: TDOMNode;
    i:integer;
begin
{ !!! Данный алгоритм не учитывает создание автоматических стилей в секции
  office:automatic-styles, т.е. при создании новой колонки её невидно из-за
  нулевой ширины !!! }

  // создадим новую ноду колонки
  Node:=FDocument.CreateElement('table:table-column');

  // попробуем найти первую строку
  FirstRowNode:=RootNode.FindNode('table:table-row');
  // если нашлась
  if Assigned(FirstRowNode) then RootNode.InsertBefore(Node,FirstRowNode)
  // если строки отсутствуют
  else RootNode.AppendChild(Node);

  {добавляем новые ячейки в каждую строку}
  for i:=0 to RootNode.ChildNodes.Count-1 do
    if RootNode.ChildNodes.Item[i].NodeName='table:table-row' then
      begin
        Node:=FDocument.CreateElement('table:table-cell');
          TDOMElement(Node).SetAttribute('table:style-name',Name+DefaultCellStyle);
          TDOMElement(Node).SetAttribute('office:value-type','string');
        Root:=RootNode.ChildNodes.Item[i].AppendChild(Node);
        Node:=FDocument.CreateElement('text:p');
          if length(TextStyle)>0 then
            TDOMElement(Node).SetAttribute('text:style-name',TextStyle)
          else
            TDOMElement(Node).SetAttribute('text:style-name',DefTextStyle);
        Root.AppendChild(Node);
      end;
end;

procedure TOdtTable.AppendRow(TextStyle: string);
var Root, Node: TDOMNode;
    i:integer;
begin
  Node:=FDocument.CreateElement('table:table-row');
  Root:=RootNode.AppendChild(Node);//Root = table:table-row
  for i:=1 to ColCount do
    begin
      Node:=FDocument.CreateElement('table:table-cell');
        TDOMElement(Node).SetAttribute('table:style-name',Name+DefaultCellStyle);
        TDOMElement(Node).SetAttribute('office:value-type','string');
      Root:=Root.AppendChild(Node);//root = table:table-cell
      Node:=FDocument.CreateElement('text:p');
        if length(TextStyle)>0 then
          TDOMElement(Node).SetAttribute('text:style-name',TextStyle)
        else
          TDOMElement(Node).SetAttribute('text:style-name',DefTextStyle);
      Root.AppendChild(Node);
      Root:=Root.ParentNode;//Root = table:table-row
    end;
end;

procedure TOdtTable.InsertRow;
var Root, Node,PreviousRoot,PreviousCol,PreviousText: TDOMNode;
    i,j:integer;
begin
  Node:=FDocument.CreateElement('table:table-row');
  PreviousRoot := RootNode.LastChild;
  TDOMElement(Node).SetAttribute('table:style-name',TDOMElement(PreviousRoot).AttribStrings['table:style-name']);
  Root:=RootNode.AppendChild(Node);//Root = table:table-row

  for i:=0 to ColCount-1 do
    begin
      Node:=FDocument.CreateElement('table:table-cell');

      //Ищем стиль ячейки от предедущей ячейки
      if (i <= PreviousRoot.ChildNodes.Count-1)then
        PreviousCol:=PreviousRoot.ChildNodes.Item[i]
        else
          PreviousCol:=nil;

      if assigned(PreviousCol)
        then
         TDOMElement(Node).SetAttribute('table:style-name',
                   TDOMElement(PreviousCol).AttribStrings['table:style-name'])
         else
         TDOMElement(Node).SetAttribute('table:style-name',Name+DefaultCellStyle) ;

        TDOMElement(Node).SetAttribute('office:value-type','string');

      Root:=Root.AppendChild(Node);//root = table:table-cell
      Node:=FDocument.CreateElement('text:p');

      //Ищем Стиль текста от предедущей строки
      if assigned(PreviousCol)  and assigned(PreviousCol.FirstChild)then
      begin

        for j:=0 to PreviousCol.ChildNodes.Count-1 do
        begin
          PreviousText:=PreviousCol.ChildNodes.Item[j];
          if PreviousText.NodeName='text:p' then break;
        end;
        if (j=PreviousCol.ChildNodes.Count-1)and
           (PreviousText.NodeName<>'text:p') then
        PreviousText:=nil;
      end
        else
          PreviousText:=nil;

      if assigned(PreviousText)
        then
        TDOMElement(Node).SetAttribute('text:style-name',
                   TDOMElement(PreviousText).AttribStrings['text:style-name'])
         else
         TDOMElement(Node).SetAttribute('text:style-name',DefTextStyle);// else

      Root.AppendChild(Node);
      Root:=Root.ParentNode;//Root = table:table-row
    end;
end;

// удаление пустых строк из таблицы
procedure TOdtTable.RemoveEmptyRow(Prefix: string);
var Node, SuicideNode: TDOMNode;
begin
  Node:=RootNode.FirstChild;
  while Assigned(Node) do begin
    //обработаем объединенные ячейки
    if Node.NodeName='table:covered-table-cell' then begin
      Node:=Node.NextSibling;
      continue;
    end;
    if Node.HasChildNodes then
      if (Node.FirstChild.NodeName='text:p') and
        (UTF8Pos(Prefix,Node.FirstChild.TextContent)>0) then begin
        SuicideNode:=Node;
        Node:=Node.NextSibling;
        SuicideNode.Destroy;
      end;
  end;
end;

procedure TOdtTable.RemoveRow;
begin
RootNode.RemoveChild(RootNode.LastChild);//последнюю строку
end;

procedure TOdtTable.RemoveRow(Count: Integer);
var i:integer;
begin
  i:=0;
  repeat
    RemoveRow;
    inc(i);
  until I=Count;
end;

procedure TOdtTable.RemoveColumn;
var i:integer;
    List: TDOMNodeList;
begin
  List:=RootNode.ChildNodes;
  for i:=1 to List.Count-1 do
    //в каждой из строк удаляем последнюю ячейку
    List.Item[i].RemoveChild(List.Item[i].LastChild);
TDOMElement(RootNode.FirstChild).AttribStrings['table:number-columns-repeated']:=IntToStr(ColCount-1);
end;

function TOdtTable.CheckFontName(AFontName: string): boolean;
var List: TDOMNodeList;
    i:integer;
begin
  List:=FDocument.DocumentElement.GetElementsByTagName('style:font-face');
  for i:=0 to List.Count-1 do
    begin
      if UpperCase(TDOMElement(List.Item[i]).AttribStrings['style:name'])=UpperCase(AFontName) then
        begin
          Result:=true;
          break;
        end;
    end;
end;

procedure TOdtTable.InsertNewFont(AFontName: string);
var Root,Node: TDOMNode;
begin
  Root:=FDocument.DocumentElement.FindNode('office:font-face-decls');
  Node:=FDocument.CreateElement('style:font-face');
    TDOMElement(Node).SetAttribute('style:name',AFontName);
    TDOMElement(Node).SetAttribute('svg:font-family',AFontName);
    TDOMElement(Node).SetAttribute('style:font-family-generic','roman');
  Root.AppendChild(Node);
end;

procedure TOdtTable.SetName(aName: string);
begin
  if aName<>'' then
  begin
    TDOMElement(FRoot).AttribStrings['table:name']:=aName;
    FName := aName;
  end;
end;


{ TOds }

constructor TOds.Create;
begin
  inherited;
end;

destructor TOds.Destroy;
begin
  inherited;
end;

procedure TOds.GenerateManifest;
var Root, Parent: TDOMNode;
begin
  inherited;
  // после выполнения стандартной процедуры генерации
  // метод добавляет элементы
  // специфичные для электронных таблиц
  Root := FManifest.GetElementsByTagName('manifest:manifest').Item[0];
  Parent:=FManifest.CreateElement('manifest:file-entry');
  TDOMElement(Parent).SetAttribute('manifest:version','1.2');
  TDOMElement(Parent).SetAttribute('manifest:full-path','/');
  TDOMElement(Parent).SetAttribute('manifest:media-type', 'application/vnd.oasis.opendocument.spreadsheet');
  Root.AppendChild(Parent);
end;

procedure TOds.GenerateContent;
var Root, Parent: TDOMNode;
begin
  // после выполнения стандартной процедуры генерации
  // метод добавляет элементы
  // специфичные для электронных таблиц
  inherited;
  Root:=FContent.DocumentElement.FindNode('office:body');
  Parent:=FContent.CreateElement('office:spreadsheet');
  Root:=Root.AppendChild(Parent);

  Parent:=FContent.CreateElement('table:table');
  Root.AppendChild(Parent);
end;

function TOds.LoadFromFile(FileName: string): boolean;
begin
  // после выполнения стандартной процедуры загрузки
  // метод подгружает корневой элемент документа office:spreadsheet
  // специфичный для электронных таблиц
  Result := inherited LoadFromFile(FileName);
  if Result then
    FRoot := FContent.GetElementsByTagName('office:spreadsheet').Item[0];
end;

procedure TOds.GenerateDocument(DocumentName: string = DefaultOdsFileName;
                                const DocumentPath: string = 'default');
begin
  // метод создан для задания DocumentName по-умолчанию
  // ВНИМАНИЕ!!! порядок параметров изменён!
  inherited;
end;

procedure TOds.ShowDocument(DocumentName: string = DefaultOdsFileName;
                                     Editor: string = 'default');
begin
  // метод создан для задания DocumentName по-умолчанию
  // ВНИМАНИЕ!!! порядок параметров изменён!
  inherited;
end;

function TOds.PrintDocument(DocumentName: string = DefaultOdsFileName): boolean;
begin
  // метод создан для задания DocumentName по-умолчанию
  Result := inherited PrintDocument(DocumentName);
end;

function TOds.GetSheet(SheetName: string): TOdsSheet;
begin
  // метод создан для инкапсуляции стандартного метода GetTable
  // в удобном для понимания виде
  Result := TOdsSheet(inherited GetTable(SheetName));
end;

function TOds.ConvertTo(InFileName: string; ext: string; var OutFileName: string
  ): boolean;
begin
  Result := inherited ConvertTo(InFileName,ext,OutFileName);
end;

function TOds.ConvertTo(InFileName: string; ext: TSupportedExtensions;
  var OutFileName: string): boolean;
begin
  Result := inherited ConvertTo(InFileName,ext,OutFileName);
end;

{ TOdf }

function TOdf.GetMetaAuthor: string;
begin
  Result:='';
  if not Assigned(FMeta) then Exit;
  if FMeta.GetElementsByTagName('meta:initial-creator').Length>0 then
    Result:=FMeta.GetElementsByTagName('meta:initial-creator').Item[0].TextContent;
end;

function TOdf.GetMetaGenerator: string;
begin
  Result:='';
  if not Assigned(FMeta) then Exit;
  if FMeta.GetElementsByTagName('meta:generator').Length>0 then
    Result:=FMeta.GetElementsByTagName('meta:generator').Item[0].TextContent;
end;

procedure TOdf.SetMetaAuthor(const Author: string);
var Node, Txt: TDOMNode;
    List: TDOMNode;
begin
  if FMeta.ChildNodes.Count=0 then GenerateMeta;
  Node:=FMeta.CreateElement('meta:initial-creator');
  Txt:=FMeta.CreateTextNode(Author);
  Node.AppendChild(Txt);
  List:=FMeta.FindNode('office:meta');
  List:=FMeta.DocumentElement;
  List.ChildNodes.Item[0].AppendChild(Node);
end;

procedure TOdf.SetMetaGenerator(const AGenerator: string);
var Node, Txt: TDOMNode;
    List: TDOMNode;
begin
  if FMeta.ChildNodes.Count=0 then GenerateMeta;
  Node:=FMeta.CreateElement('meta:generator');
  Txt:=FMeta.CreateTextNode(AGenerator);
  Node.AppendChild(Txt);
  List:=FMeta.FindNode('office:meta');
  List:=FMeta.DocumentElement;
  List.ChildNodes.Item[0].AppendChild(Node);
end;

constructor TOdf.Create;
begin
  inherited Create;
  // создаем необходимые документы
  GenerateStyles;
  GenerateContent;
  GenerateManifest;
  GenerateMeta;
  // создаем папку для временных файлов
  repeat
    TempDir := SysUtils.GetTempDir+IncludeTrailingPathDelimiter(createUniqueString);
  until not DirectoryExists(TempDir);
  CreateDir(TempDir);
end;

destructor TOdf.Destroy;
begin
  FStyles.Free;
  FContent.Free;
  FManifest.Free;
  FMeta.Free;
  DeleteDirectory(TempDir,False);
end;

procedure TOdf.InsertXMLNS(var RootNode: TDOMElement);
var i:integer;
begin
  for i:=1 to High(xmlns) do
    RootNode.SetAttribute('xmlns:'+xmlns[i,1],xmlns[i,2]);
end;

procedure TOdf.GenerateMeta;

  function MakeDate(const ADate: TDateTime): string;
  var AYear, AMonth, ADay, AHour, AMinute, ASecond, AMilliSecond: Word;
  begin
    DecodeDateTime(ADate, AYear, AMonth, ADay, AHour, AMinute, ASecond, AMilliSecond);
    Result:=IntToStr(AYear)+'-';

    if AMonth<10 then
      Result:=Result+'0'+IntToStr(AMonth)+'-'
    else
      Result:=Result+IntToStr(AMonth)+'-';

    if ADay<10 then
      Result:=Result+'0'+IntToStr(ADay)
    else
      Result:=Result+IntToStr(ADay);
    Result:=Result+'T'+IntToStr(AHour)+':'+IntToStr(AMinute)+':'+IntToStr(ASecond);
  end;

var Root,Parent,Txt: TDOMNode;
begin
  if Assigned(FMeta) then FreeAndNil(FMeta);
  FMeta:=TXMLDocument.Create;
  Root:=FMeta.CreateElement('office:document-meta');
  TDOMElement(Root).SetAttribute('xmlns:office','urn:oasis:names:tc:opendocument:xmlns:office:1.0');
  TDOMElement(Root).SetAttribute('xmlns:meta','urn:oasis:names:tc:opendocument:xmlns:meta:1.0');
  FMeta.Appendchild(Root);
  Root:=FMeta.DocumentElement;
  Parent:=FMeta.CreateElement('office:meta');
  Root.AppendChild(Parent);
  Parent:=FMeta.DocumentElement;
  Root:=FMeta.CreateElement('meta:creation-date');
  Txt:=FMeta.CreateTextNode(MakeDate(Now));
  Root.AppendChild(Txt);
  Parent.ChildNodes.Item[0].AppendChild(Root);
end;

procedure TOdf.GenerateManifest;
var Root,Parent: TDOMNode;
    i:integer;
begin
 if Assigned(FManifest) then FreeAndNil(FManifest);
 FManifest:=TXMLDocument.Create;

 Root:=FManifest.CreateElement('manifest:manifest');
 TDOMElement(Root).SetAttribute('xmlns:manifest','urn:oasis:names:tc:opendocument:xmlns:manifest:1.0');
 FManifest.Appendchild(Root);
 for i:=1 to 3 do
   begin
     Parent:=FManifest.CreateElement('manifest:file-entry');
     TDOMElement(Parent).SetAttribute('manifest:full-path',DocEntrys[i]);
     TDOMElement(Parent).SetAttribute('manifest:media-type','text/xml');
     Root.AppendChild(Parent);
    end;
end;

procedure TOdf.GenerateContent;
var Root,Parent: TDOMNode;
begin
  FreeAndNil(FContent);
  FContent:=TXMLDocument.Create;
  Root:=FContent.CreateElement('office:document-content');
    InsertXMLNS(TDOMElement(Root));
    TDOMElement(Root).SetAttribute('office:version','1.2');
  Root:=FContent.AppendChild(Root);
  Root:=FContent.DocumentElement;
  Parent:=FContent.CreateElement('office:scripts');
  Root.AppendChild(Parent);
  Parent:=FContent.CreateElement('office:font-face-decls');
  Root:=Root.AppendChild(Parent);
  Parent:=FContent.CreateElement('style:font-face');
    TDOMElement(Parent).SetAttribute('style:name','Arial');
    TDOMElement(Parent).SetAttribute('svg:font-family','Arial');
  Root.AppendChild(Parent);
  Root:=FContent.DocumentElement;
  Parent:=FContent.CreateElement('office:automatic-styles');
  Root.AppendChild(Parent);

  Parent:=FContent.CreateElement('office:body');
  Root.AppendChild(Parent);
end;

procedure TOdf.GenerateStyles;
var Root, Node: TDOMNode;
begin
FreeAndNil(FStyles);
FStyles:=TXMLDocument.Create;
Root:=FStyles.CreateElement('office:document-styles');
  InsertXMLNS(TDOMElement(Root));
  TDOMElement(Root).SetAttribute('office:version','1.2');
Root:=FStyles.AppendChild(Root); //текущий узел office:document-styles
Node:=FStyles.CreateElement('office:font-face-decls');
Root:=Root.AppendChild(Node); //текущий узел office:font-face-decls
Node:=FStyles.CreateElement('style:font-face');
  TDOMElement(Node).SetAttribute('style:name','Arial');
  TDOMElement(Node).SetAttribute('svg:font-family','Arial');
Root.AppendChild(Node);
Root:=FStyles.DocumentElement; //текущий узел office:document-styles
Node:=FStyles.CreateElement('office:styles');
Root:=Root.AppendChild(Node);  //текущий узел office:styles
Node:=Fstyles.CreateElement('style:default-style');
 TDOMElement(Node).SetAttribute('style:family','graphic');
Root:=Root.AppendChild(Node);  //текущий узел style:default-style
Node:=FStyles.CreateElement('style:graphic-properties');
  TDOMElement(Node).SetAttribute('draw:shadow-offset-x','0.3cm');
  TDOMElement(Node).SetAttribute('draw:shadow-offset-y','0.3cm');
  TDOMElement(Node).SetAttribute('draw:start-line-spacing-horizontal','0.283cm');
  TDOMElement(Node).SetAttribute('draw:start-line-spacing-vertical','0.283cm');
  TDOMElement(Node).SetAttribute('draw:end-line-spacing-horizontal','0.283cm');
  TDOMElement(Node).SetAttribute('draw:end-line-spacing-vertical','0.283cm');
  TDOMElement(Node).SetAttribute('style:flow-with-text','false');
Root.AppendChild(Node);  //текущий узел style:default-style
Node:=FStyles.CreateElement('style:paragraph-properties');
TDOMElement(Node).SetAttribute('style:text-autospace','ideograph-alpha');
  TDOMElement(Node).SetAttribute('style:line-break','strict');
  TDOMElement(Node).SetAttribute('style:writing-mode','lr-tb');
  TDOMElement(Node).SetAttribute('style:font-independent-line-spacing','false');
Root:=Root.AppendChild(Node); //текущий узел style:default-style
Node:=FStyles.CreateElement('style:tab-stops');
Root:=Root.AppendChild(Node).ParentNode.ParentNode; //текущий узел style:default-style
Node:=FStyles.CreateElement('style:text-properties');
  TDOMElement(Node).SetAttribute('style:use-window-font-color','true');
  TDOMElement(Node).SetAttribute('fo:font-size','10pt');
  TDOMElement(Node).SetAttribute('fo:language','ru');
  TDOMElement(Node).SetAttribute('fo:country','RU');
  TDOMElement(Node).SetAttribute('style:letter-kerning','true');
Root:=Root.AppendChild(Node).ParentNode.ParentNode; //текущий узел office:styles
Node:=FStyles.CreateElement('style:default-style');
  TDOMElement(Node).SetAttribute('style:family','paragraph');
Root:=Root.AppendChild(Node); //style:default-style
Node:=FStyles.CreateElement('style:paragraph-properties');
  TDOMElement(Node).SetAttribute('fo:hyphenation-ladder-count','no-limit');
  TDOMElement(Node).SetAttribute('style:text-autospace','ideograph-alpha');
  TDOMElement(Node).SetAttribute('style:punctuation-wrap','hanging');
  TDOMElement(Node).SetAttribute('style:line-break','strict');
  TDOMElement(Node).SetAttribute('style:tab-stop-distance','1.251cm');
  TDOMElement(Node).SetAttribute('style:writing-mode','page');
Root.AppendChild(Node);//style:default-style
Node:=FStyles.CreateElement('style:text-properties');
  TDOMElement(Node).SetAttribute('style:use-window-font-color','true');
  TDOMElement(Node).SetAttribute('style:font-name','Arial');
  TDOMElement(Node).SetAttribute('fo:font-size','10pt');
  TDOMElement(Node).SetAttribute('fo:language','ru');
  TDOMElement(Node).SetAttribute('fo:country','RU');
  TDOMElement(Node).SetAttribute('style:letter-kerning','true');
Root:=Root.AppendChild(Node).ParentNode.ParentNode; //текущий узел office:styles
Node:=FStyles.CreateElement('style:default-style');
  TDOMElement(Node).SetAttribute('style:family','table');
Root:=Root.AppendChild(Node); //текущий узел style:default-style
Node:=FStyles.CreateElement('style:table-properties');
  TDOMElement(Node).SetAttribute('table:border-model','collapsing');
Root:=Root.AppendChild(Node).ParentNode.ParentNode; //текущий узел office:styles
Node:=FStyles.CreateElement('style:default-style');
  TDOMElement(Node).SetAttribute('style:family','table-row');
Root:=Root.AppendChild(Node); //текущий узел style:default-style
Node:=FStyles.CreateElement('style:table-row-properties');
  TDOMElement(Node).SetAttribute('fo:keep-together','auto');
Root:=Root.AppendChild(Node).ParentNode.ParentNode; //текущий узел office:styles
Node:=FStyles.CreateElement('style:style');
  TDOMElement(Node).SetAttribute('style:name','Standard');
  TDOMElement(Node).SetAttribute('style:family','paragraph');
  TDOMElement(Node).SetAttribute('style:class','text');
Root.AppendChild(Node); //текущий узел office:styles
Node:=FStyles.CreateElement('style:style');
  TDOMElement(Node).SetAttribute('style:name','Heading');
  TDOMElement(Node).SetAttribute('style:family','paragraph');
  TDOMElement(Node).SetAttribute('style:parent-style-name','Standard');
  TDOMElement(Node).SetAttribute('style:next-style-name','Text_20_body');
  TDOMElement(Node).SetAttribute('style:class','text');
Root:=Root.AppendChild(Node);//текущий узел style:style
Node:=FStyles.CreateElement('style:paragraph-properties');
  TDOMElement(Node).SetAttribute('fo:margin-top','0.423cm');
  TDOMElement(Node).SetAttribute('fo:margin-bottom','0.212cm');
  TDOMElement(Node).SetAttribute('fo:keep-with-next','always');
Root.AppendChild(Node);//текущий узел style:style
Node:=FStyles.CreateElement('style:text-properties');
  TDOMElement(Node).SetAttribute('style:font-name','Arial');
  TDOMElement(Node).SetAttribute('fo:font-size','14pt');
Root:=Root.AppendChild(Node).ParentNode.ParentNode; //текущий узел office:styles
Node:=FStyles.CreateElement('style:style');
  TDOMElement(Node).SetAttribute('style:name','Text_20_body');
  TDOMElement(Node).SetAttribute('style:display-name','Text body');
  TDOMElement(Node).SetAttribute('style:family','paragraph');
  TDOMElement(Node).SetAttribute('style:parent-style-name','Standard');
  TDOMElement(Node).SetAttribute('style:class','text');
Root:=Root.AppendChild(Node);//текущий узел style:style
Node:=FStyles.CreateElement('style:paragraph-properties');
  TDOMElement(Node).SetAttribute('fo:margin-top','0cm');
  TDOMElement(Node).SetAttribute('fo:margin-bottom','0.212cm');
Root.AppendChild(Node);
{ ниже не изменять }
//Root:=FStyles.DocumentElement;//текущий узел office:document-styles
//WriteXMLFile(FStyles,ExtractFilePath(Application.ExeName)+'Test.xml');
end;

function TOdf.DocumentLoaded: boolean;
begin
  Result :=  ((Assigned(FManifest)) and
             (FManifest.ChildNodes.Count<>0) and (FName<>''))
end;

function TOdf.CloseDocument: boolean;
begin
  if DocumentLoaded then
  begin
    FreeAndNil(FContent);
    FreeAndNil(FStyles);
    FreeAndNil(FMeta);
    FreeAndNil(FManifest);
    FreeAndNil(FSettings);
  end;
  FName :='';
  Result:=DeleteDirectory(TempDir, true);
end;

function TOdf.LoadPartOfDocument(FileName: string; Doc: TFileType): boolean;
var f: TXMLDocument;
begin
  try
    ReadXMLFile(f, FileName);
    case Doc of
      ftStyles   : Styles:=f;
      ftContent  : Content:=f;
      ftManifest : Manifest:=f;
      ftMeta     : Meta:=f;
      ftSettings : Settings:=f;
    end;
    Result:=true;
  except
    Result:=false;
  end;
end;

function TOdf.LoadFromFile(FileName: string): boolean;
var
    UnZipper: TUnZipper;
begin
  if FileExistsUTF8(FileName) then
    try
      CloseDocument;
      try
        //распаковываем шаблон
        UnZipper := TUnZipper.Create;
        UnZipper.FileName := FileName;
        UnZipper.OutputPath := TempDir;
        UnZipper.Examine;
        UnZipper.UnZipAllFiles;

        //подгружаем данные
        LoadPartOfDocument(TempDir+'styles.xml',ftStyles);
        LoadPartOfDocument(TempDir+'content.xml',ftContent);
        LoadPartOfDocument(TempDir+IncludeTrailingPathDelimiter('META-INF')+'manifest.xml',ftManifest);
        LoadPartOfDocument(TempDir+'meta.xml',ftMeta);
        LoadPartOfDocument(TempDir+'settings.xml',ftSettings);

        //устанавливаем
        FName:=ChangeFileExt(ExtractFileName(FileName),'');
        Result:=true;
      except
        Result:=false;
      end;
    finally
       UnZipper.Free;
    end
  else
    Result := false;
end;

function TOdf.GetTable(TableName: string): TOdfTable;
begin
  Result := TOdfTable.Create(FContent,TableName);
  if not Assigned(Result) then
    MessageDlg('Ошибка','Таблица с названием "'+TableName+'" не найдена в документе',mtError,[mbOK],0);
end;

function TOdf.FindAndReplace(Search, Replace: string):boolean;
begin
  try
    ReplaceTextInChildNodes(FRoot, Search, Replace);
    Result:=true;
  except
    Result:=false;
  end;
end;

// Удаление ноды содержащей определенный текст
// ВНИМАНИЕ!!!
// Сомнительная функция - при неосмотрительном использовании структура документа
// может быть нарушена!
// Для простого удаления текста лучше использовать FindAndReplace, заменяя
// текст для поиска пустой строкой
function TOdf.FindAndRemove(Search: string):boolean;
begin
  try
    RemoveNodesWithText(FRoot, Search);
    Result:=true;
  except
    Result:=false;
  end;
end;

// soffice не работает в LO и ООо с версии 4... имхо надо что-то другое (c)Leo
function TOdf.GetOfficeVersion: string;
//var
   //AProcess: TProcess;
   //AStringList: TStringList;
   //i: integer;
 // Начинаем нашу программу
 begin
   //// Создаем объект TProcess
   //AProcess := TProcess.Create(nil);
   //
   //// Создаем объект TStringList
   //AStringList := TStringList.Create;
   //
   //// Зададим командную строку
   //AProcess.CommandLine := 'soffice --version';
   //
   //// Установим опции программы. Первая из них не позволит нашей программе
   //// выполнятся до тех пор, пока не закончит выполнение запущенная программа
   //// Также добавим опцию, которая говорит, что мы хотим прочитать
   //// вывод запущенной программы
   //AProcess.Options := AProcess.Options + [poWaitOnExit, poUsePipes];
   //
   //// Теперь запускаем программу
   //AProcess.Execute;
   //
   //// Пока запущенная программа не закончится, досюда мы не дойдем
   //
   //// А теперь прочитаем вывод в список строк TStringList.
   //AStringList.LoadFromStream(AProcess.Output);
   //
   //for i:=0 to AStringList.Count-1 do
   //begin
   //  if UTF8pos('LibreOffice',AStringList[i])>0 then
   //  Result:= AStringList[i];
   //  if UTF8pos('OpenOffice',AStringList[i])>0 then
   //  Result:= AStringList[i];
   //  if Result<>'' then break;
   //end;
   //
   //// Можем уничтожить
   //// TStringList и TProcess.
   //AStringList.Free;
   //AProcess.Free;
  Result := 'LibreOffice';
end;

// вывод документа на принтер
function TOdf.PrintDocument(FileName: string): boolean;
var AProcess: TProcess;
    office: string;
begin
  Result := false;
  if FileExists(FileName) and (GetOfficeVersion<>'') then
     begin
        //Выполним команду печати
        AProcess := TProcess.Create(nil);
        AProcess.Executable := 'libreoffice4.0';
        AProcess.Parameters.Add('--invisible');
        AProcess.Parameters.Add('-p "'+FileName+'"');
        AProcess.Options := AProcess.Options + [poWaitOnExit];
        AProcess.Execute;
        AProcess.Free;
        Result := true;
     end else
         Result:=false;
end;

// просмотр получившегося документа
procedure TOdf.ShowDocument(DocumentName: string;
                            Editor: string = 'default');
var Proc: TProcess;
begin
  if Editor = '' then Editor := 'default';
  GenerateDocument(DocumentName, TempDir);
  if Editor = 'default' then
    begin
      if not OpenDocument(TempDir + DocumentName) then
        MessageDlg('Ошибка','Не удалось открыть файл "' +
                                TempDir + DocumentName + '"',mtError,[mbOK],0);
      sleep(3000);
    end
  else
    try
      Proc := TProcess.Create(nil);
      Proc.Executable := Editor;
      Proc.Parameters.Append(TempDir + DocumentName);
      Proc.Options := [poWaitOnExit];
      Proc.ShowWindow := swoShowMaximized;
      Proc.Execute;
    finally
      if Proc.WaitOnExit then Proc.Free;
    end;
end;

// генерация документа
procedure TOdf.GenerateDocument(DocumentName: string;
                          const DocumentPath: string = 'default');
var Zipper: TZipper;

  // добавим в архив всё что есть в директории, включая папки с файлами
  procedure AddEntries(dir:string);
  var F:TSearchRec;
      ires:integer;
      short_dir:string;
  begin
    ires:=FindFirst(IncludeTrailingPathDelimiter(dir)+'*',faAnyFile,F);
    while ires=0 do begin
      if (F.Name<>'.') and (F.Name<>'..') then
        if (F.Attr and faDirectory > 0) then
          AddEntries(IncludeTrailingPathDelimiter(dir)+F.Name)
        else begin
          if dir<>TempDir then begin
            short_dir:=UTF8Copy(dir,UTF8Length(TempDir)+1,UTF8Length(dir)-UTF8Length(TempDir));
            Zipper.Entries.AddFileEntry(IncludeTrailingPathDelimiter(dir)+
              F.Name,IncludeTrailingPathDelimiter(short_dir)+F.Name);
          end
          else
            Zipper.Entries.AddFileEntry(IncludeTrailingPathDelimiter(dir)+
              F.Name,F.Name);
        end;
      ires:=FindNext(F);
    end;
    FindClose(F);
  end;

begin

  if (Not Assigned(FManifest)) or (FManifest.ChildNodes.Count=0) then
    GenerateManifest;

  if not DirectoryExists(DocumentPath) and (DocumentPath<>'default') then
    CreateDir(DocumentPath);

  //удалим старые xml-файлы
  DeleteFile(TempDir+'content.xml');
  DeleteFile(TempDir+'styles.xml');
  DeleteFile(TempDir+'meta.xml');
  DeleteFile(TempDir+'settings.xml');
  DeleteFile(IncludeTrailingPathDelimiter(TempDir+'META-INF')+'manifest.xml');

  //сохраняем новые файлы
  if not DirectoryExists(TempDir+'META-INF') then CreateDir(TempDir+'META-INF');
  WriteXMLFile(FContent, TempDir+'content.xml');
  WriteXMLFile(FStyles, TempDir+'styles.xml');
  WriteXMLFile(FMeta, TempDir+'meta.xml');
  WriteXMLFile(FManifest,IncludeTrailingPathDelimiter(TempDir+'META-INF')+'manifest.xml');
  WriteXMLFile(FSettings, TempDir+'settings.xml');

  try
    Zipper := TZipper.Create;
    if DocumentPath = 'default' then
      Zipper.FileName := ExtractFilePath(Application.ExeName) + DocumentName
    else
      Zipper.FileName := IncludeTrailingPathDelimiter(DocumentPath) + DocumentName;
    AddEntries(TempDir);
    Zipper.ZipAllFiles;
  finally
    Zipper.Free;
  end;
end;

function TOdf.ConvertTo(InFileName: string; to_ext_str: string; var OutFileName: string): boolean;
var i:integer;
begin
  for i:=0 to Length(SupportedExtensionsArr) do
    if SupportedExtensionsArr[i]=to_ext_str then
      Result:=ConvertTo(InFileName,TSupportedExtensions(i),OutFileName);
end;

function TOdf.ConvertTo(InFileName: string; to_ext: TSupportedExtensions; var OutFileName: string): boolean;
var AProcess: TProcess;
    office: string;
    from_ext: TSupportedExtensions;
    from_ext_str, to_ext_str: string;
begin
  Result := false;

  case to_ext of
    seODT: from_ext:=seDOC;
    seODS: from_ext:=seXLS;
    sePDF, seDOC: from_ext:=seODT;
    seXLS: from_ext:=seODS;
  end;

  from_ext_str := SupportedExtensionsArr[Ord(from_ext)];
  to_ext_str := SupportedExtensionsArr[Ord(to_ext)];

  OutFileName := UTF8Copy(InFileName,1,UTF8Pos(from_ext_str,InFileName)-1) + to_ext_str;

  if FileExists(InFileName) and (GetOfficeVersion<>'') and
    (LowerCase(ExtractFileExt(InFileName)) = SupportedExtensionsArr[Ord(from_ext)]) and
    ExportScriptToOO() then

    begin
      //Выполним команду конверсии документа
      AProcess := TProcess.Create(nil);
      case to_ext of
        sePDF: AProcess.CommandLine :=
                'libreoffice4.0 --invisible "macro:///Standard.macrosODF.SaveAsPDF('+
                InFileName+')"';
        seDOC: AProcess.CommandLine :=
                'libreoffice4.0 --invisible "macro:///Standard.macrosODF.SaveAsDoc('+
                InFileName+')"';
        seODT: AProcess.CommandLine :=
                'libreoffice4.0 --invisible "macro:///Standard.macrosODF.SaveAsOOO('+
                InFileName+')"';
        seXLS: AProcess.CommandLine :=
                'libreoffice4.0 --invisible "macro:///Standard.macrosODF.SaveAsXls('+
                InFileName+')"';
      end;
      AProcess.Options := AProcess.Options + [poWaitOnExit];
      AProcess.Execute;
      AProcess.Free;
      Result := true;
    end
  else Result:=false;
end;

function TOdf.ExportScriptToOO(): boolean;
var s,v: UTF8string;
    ADoc: TXMLDocument;
    Root,node:TDOMNode;
begin
  Result := false;
  //Найдем домашний каталог пользователя
  s:=GetEnvironmentVariable('HOME');
  if s='/root' then
  begin
    Showmessage('Программа запущена из под Root. Команда ConvertTo не выполнена');
    Exit;
  end;
  v:=GetOfficeVersion();
  //найдем каталог офиса
  if DirectoryExists(s+IncludeTrailingPathDelimiter('/.config/libreoffice/4/'))and
    (pos('LibreOffice',v)>0) then
  begin
    s:=s+IncludeTrailingPathDelimiter('/.config/libreoffice/4/');
  end else
    if DirectoryExists(s+IncludeTrailingPathDelimiter('/.config/openoffice/4/'))and
      (pos('OpenOffice',v)>0) then
    begin
      s:=s+IncludeTrailingPathDelimiter('/.config/openoffice/4/');
    end
      else
        s:='';

  if s='' then Exit;
  s:=s+IncludeTrailingPathDelimiter('user/basic/Standard/');
  //Cкопируем скрипт в каталог офиса(Если конечно скрипт сущществует)
  if fileExists('macrosODF.xba') then
  begin
    //Проверим в каталоге назначения есть ли скрипт
    if not fileExists(s + 'macrosODF.xba') then
      Result := CopyFile('macrosODF.xba',s + 'macrosODF.xba')
      else Result := true;
  end
   else
     Result:=false;

   if not Result then Exit;

   //Зарегестрируем\проверим есть ли такой макрос
   if FileExists(s+'script.xlb') then
   begin
      ReadXMLFile(ADoc,s+'script.xlb');
     //  <library:element library:name="macros"/>
      root:=ADoc.GetElementsByTagName('library:library').Item[0];
      node:=root.FirstChild;
      while node<>nil do
      begin
         if TDOMElement(node).AttribStrings['library:name']='macrosODF' then break;
         node:=node.NextSibling;
      end;
      if node=nil then
      begin
        node:=ADoc.CreateElement('library:element');
        TDOMElement(node).SetAttribute('library:name','macrosODF');
        root.AppendChild(node);
      end;
      WriteXMLFile(ADoc,s+'script.xlb');
      Result := true;
   end else
       Result:=false;
end;

end.

