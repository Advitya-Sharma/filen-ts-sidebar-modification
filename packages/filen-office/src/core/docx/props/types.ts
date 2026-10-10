import type { XmlDocument } from "../../xml/index"

// Typed WordprocessingML properties (ECMA-376 Part 1 §17.3, §17.4, §17.6), in the units OOXML stores:
//   twips       1/20 pt (indents, spacing, page geometry, table widths of type "dxa")
//   halfPoints  1/2 pt (font sizes, kerning threshold, vertical position)
//   eighths     1/8 pt (border widths)
//   fiftieths   1/50 % (widths of type "pct")
// Direct property objects are sparse (a field is present only when the XML sets it) and interned by their
// source bytes, so equal XML shares one frozen object. Resolved objects have every field and are interned
// by value.

// A node of a parsed part: how the model refers to raw property XML it does not interpret.
export interface XmlNodeRef {
	readonly document: XmlDocument
	readonly node: number
}

// ST_ThemeColor (§17.18.97).
export type ThemeColorName =
	| "dark1"
	| "light1"
	| "dark2"
	| "light2"
	| "accent1"
	| "accent2"
	| "accent3"
	| "accent4"
	| "accent5"
	| "accent6"
	| "hyperlink"
	| "followedHyperlink"
	| "background1"
	| "text1"
	| "background2"
	| "text2"
	| "none"

// A colour as written: `val` is "auto" or six hex digits (upper case); theme attributes win over it when a
// theme is present. Tint and shade are 0..255 (ST_UcharHexNumber).
export interface ColorProps {
	readonly val?: string
	readonly themeColor?: ThemeColorName
	readonly themeTint?: number
	readonly themeShade?: number
}

// CT_Border (§17.3.4). `val` is ST_Border ("single", "double", "nil", "none", art borders...).
export interface BorderProps extends ColorProps {
	readonly style: string
	// Eighths of a point.
	readonly size?: number
	// Points.
	readonly space?: number
	readonly shadow?: boolean
	readonly frame?: boolean
}

// CT_Shd (§17.3.5). `pattern` is ST_Shd ("clear", "solid", "pct25", "horzStripe"...).
export interface ShadingProps {
	readonly pattern: string
	readonly color?: ColorProps
	readonly fill?: ColorProps
}

export type WidthType = "nil" | "pct" | "dxa" | "auto"

// CT_TblWidth (§17.4.88). `value` is twips for "dxa", fiftieths of a percent for "pct", 0 otherwise.
export interface WidthProps {
	readonly value: number
	readonly type: WidthType
}

export interface FontsProps {
	readonly ascii?: string
	readonly hAnsi?: string
	readonly eastAsia?: string
	readonly cs?: string
	readonly asciiTheme?: ThemeFontRef
	readonly hAnsiTheme?: ThemeFontRef
	readonly eastAsiaTheme?: ThemeFontRef
	readonly csTheme?: ThemeFontRef
	readonly hint?: "default" | "eastAsia" | "cs"
}

// ST_Theme (§17.18.96).
export type ThemeFontRef =
	"majorAscii" | "majorHAnsi" | "majorEastAsia" | "majorBidi" | "minorAscii" | "minorHAnsi" | "minorEastAsia" | "minorBidi"

export interface UnderlineProps extends ColorProps {
	// ST_Underline: "single", "words", "double", "thick", "dotted", "dash", "wave", "none"...
	readonly style: string
}

export interface LanguageProps {
	readonly latin?: string
	readonly eastAsia?: string
	readonly bidi?: string
}

export interface EastAsianLayoutProps {
	readonly id?: number
	readonly combine?: boolean
	readonly combineBrackets?: string
	readonly vertical?: boolean
	readonly verticalCompress?: boolean
}

export interface FitTextProps {
	// Twips the run's text is fitted into.
	readonly width: number
	readonly id?: number
}

// A tracked property change (w:rPrChange, w:pPrChange, ...): who changed it, when, and the properties as
// they were before. The element itself stays in the XML; `propertyChangeNode` finds it for the writer.
export interface PropertyChange<T> {
	readonly id?: number
	readonly author?: string
	readonly date?: string
	readonly previous: T
}

// "merge" is w:cellMerge.
export type RevisionKind = "ins" | "del" | "moveFrom" | "moveTo" | "merge"

export interface RevisionMark {
	readonly kind: RevisionKind
	readonly id?: number
	readonly author?: string
	readonly date?: string
}

// w:rPr (§17.3.2). Toggle properties (b, i, caps...) are booleans as written: false is an explicit "off".
export interface RunProps {
	readonly styleId?: string
	readonly fonts?: FontsProps
	readonly bold?: boolean
	readonly boldCs?: boolean
	readonly italic?: boolean
	readonly italicCs?: boolean
	readonly caps?: boolean
	readonly smallCaps?: boolean
	readonly strike?: boolean
	readonly doubleStrike?: boolean
	readonly outline?: boolean
	readonly shadow?: boolean
	readonly emboss?: boolean
	readonly imprint?: boolean
	readonly noProof?: boolean
	readonly snapToGrid?: boolean
	readonly vanish?: boolean
	readonly webHidden?: boolean
	readonly specVanish?: boolean
	readonly rtl?: boolean
	readonly cs?: boolean
	readonly oMath?: boolean
	readonly color?: ColorProps
	// Character spacing, twips (signed).
	readonly spacing?: number
	// Horizontal scale, percent.
	readonly scale?: number
	// Kerning threshold, half-points (0 = off).
	readonly kern?: number
	// Raised (positive) or lowered baseline, half-points.
	readonly position?: number
	readonly size?: number
	readonly sizeCs?: number
	// ST_HighlightColor name ("yellow", "darkBlue", "none"...).
	readonly highlight?: string
	readonly underline?: UnderlineProps
	readonly effect?: string
	readonly border?: BorderProps
	readonly shading?: ShadingProps
	readonly fitText?: FitTextProps
	readonly vertAlign?: "baseline" | "superscript" | "subscript"
	// ST_Em: "none", "dot", "comma", "circle", "underDot".
	readonly emphasis?: string
	readonly lang?: LanguageProps
	readonly eastAsianLayout?: EastAsianLayoutProps
	// Paragraph-mark revisions (only in w:pPr/w:rPr).
	readonly revisions?: readonly RevisionMark[]
	readonly change?: PropertyChange<RunProps>
}

export interface TabStop {
	// ST_TabJc: "clear", "start"/"left", "center", "end"/"right", "decimal", "bar", "num".
	readonly alignment: string
	// Twips (signed).
	readonly position: number
	// ST_TabTlc: "none", "dot", "hyphen", "underscore", "heavy", "middleDot".
	readonly leader?: string
}

export interface SpacingProps {
	readonly before?: number
	readonly beforeLines?: number
	readonly beforeAutospacing?: boolean
	readonly after?: number
	readonly afterLines?: number
	readonly afterAutospacing?: boolean
	// Twips for "exact"/"atLeast", 240ths of a line for "auto".
	readonly line?: number
	readonly lineRule?: "auto" | "exact" | "atLeast"
}

// Logical sides: transitional w:left/w:right are read as start/end (the 2nd edition renamed them).
// *Chars values are hundredths of a character unit and win over the twips value when non-zero.
export interface IndentProps {
	readonly start?: number
	readonly startChars?: number
	readonly end?: number
	readonly endChars?: number
	readonly hanging?: number
	readonly hangingChars?: number
	readonly firstLine?: number
	readonly firstLineChars?: number
}

export interface NumberingRef {
	readonly numId?: number
	readonly ilvl?: number
}

export interface ParagraphBordersProps {
	readonly top?: BorderProps
	readonly start?: BorderProps
	readonly bottom?: BorderProps
	readonly end?: BorderProps
	readonly between?: BorderProps
	readonly bar?: BorderProps
}

// w:framePr (§17.3.1.11): positions in twips.
export interface FrameProps {
	readonly dropCap?: "none" | "drop" | "margin"
	readonly lines?: number
	readonly width?: number
	readonly height?: number
	readonly vSpace?: number
	readonly hSpace?: number
	readonly wrap?: string
	readonly hAnchor?: string
	readonly vAnchor?: string
	readonly x?: number
	readonly xAlign?: string
	readonly y?: number
	readonly yAlign?: string
	readonly hRule?: "auto" | "exact" | "atLeast"
	readonly anchorLock?: boolean
}

// Conditional-formatting flags of w:cnfStyle (§17.3.1.8).
export interface ConditionalFlags {
	readonly firstRow?: boolean
	readonly lastRow?: boolean
	readonly firstColumn?: boolean
	readonly lastColumn?: boolean
	readonly oddVBand?: boolean
	readonly evenVBand?: boolean
	readonly oddHBand?: boolean
	readonly evenHBand?: boolean
	readonly firstRowFirstColumn?: boolean
	readonly firstRowLastColumn?: boolean
	readonly lastRowFirstColumn?: boolean
	readonly lastRowLastColumn?: boolean
}

// w:pPr (§17.3.1).
export interface ParagraphProps {
	readonly styleId?: string
	readonly keepNext?: boolean
	readonly keepLines?: boolean
	readonly pageBreakBefore?: boolean
	readonly frame?: FrameProps
	readonly widowControl?: boolean
	readonly numbering?: NumberingRef
	readonly suppressLineNumbers?: boolean
	readonly borders?: ParagraphBordersProps
	readonly shading?: ShadingProps
	readonly tabs?: readonly TabStop[]
	readonly suppressAutoHyphens?: boolean
	readonly kinsoku?: boolean
	readonly wordWrap?: boolean
	readonly overflowPunct?: boolean
	readonly topLinePunct?: boolean
	readonly autoSpaceDE?: boolean
	readonly autoSpaceDN?: boolean
	readonly bidi?: boolean
	readonly adjustRightInd?: boolean
	readonly snapToGrid?: boolean
	readonly spacing?: SpacingProps
	readonly indent?: IndentProps
	readonly contextualSpacing?: boolean
	readonly mirrorIndents?: boolean
	readonly suppressOverlap?: boolean
	// ST_Jc as written: "start"/"left", "center", "end"/"right", "both", "distribute", "*Kashida", "thaiDistribute", "numTab".
	readonly alignment?: string
	// ST_TextDirection: "lrTb"/"tb", "tbRl"/"rl", "btLr"/"lr", "lrTbV", "tbRlV", "tbLrV".
	readonly textDirection?: string
	// ST_TextAlignment: "top", "center", "baseline", "bottom", "auto".
	readonly textAlignment?: string
	readonly textboxTightWrap?: string
	// 0..8 for headings, 9 = body text.
	readonly outlineLevel?: number
	readonly divId?: number
	readonly conditional?: ConditionalFlags
	// Paragraph-mark run properties.
	readonly markRun?: RunProps
	// The paragraph carries a w:sectPr (the last paragraph of a section).
	readonly hasSection?: boolean
	readonly change?: PropertyChange<ParagraphProps>
}

export interface TableLook {
	readonly firstRow: boolean
	readonly lastRow: boolean
	readonly firstColumn: boolean
	readonly lastColumn: boolean
	readonly noHBand: boolean
	readonly noVBand: boolean
}

export interface TableBordersProps {
	readonly top?: BorderProps
	readonly start?: BorderProps
	readonly bottom?: BorderProps
	readonly end?: BorderProps
	readonly insideH?: BorderProps
	readonly insideV?: BorderProps
}

export interface CellMarginsProps {
	readonly top?: WidthProps
	readonly start?: WidthProps
	readonly bottom?: WidthProps
	readonly end?: WidthProps
}

// w:tblpPr (§17.4.58): distances and offsets in twips.
export interface TablePositionProps {
	readonly leftFromText?: number
	readonly rightFromText?: number
	readonly topFromText?: number
	readonly bottomFromText?: number
	readonly vertAnchor?: "text" | "margin" | "page"
	readonly horzAnchor?: "text" | "margin" | "page"
	readonly xSpec?: string
	readonly x?: number
	readonly ySpec?: string
	readonly y?: number
}

// w:tblPr (§17.4.60) and w:tblPrEx (§17.4.61, which carries a subset).
export interface TableProps {
	readonly styleId?: string
	readonly position?: TablePositionProps
	readonly overlap?: "never" | "overlap"
	readonly bidiVisual?: boolean
	readonly rowBandSize?: number
	readonly colBandSize?: number
	readonly width?: WidthProps
	readonly alignment?: string
	readonly cellSpacing?: WidthProps
	readonly indent?: WidthProps
	readonly borders?: TableBordersProps
	readonly shading?: ShadingProps
	readonly layout?: "fixed" | "autofit"
	readonly cellMargins?: CellMarginsProps
	readonly look?: TableLook
	readonly caption?: string
	readonly description?: string
	readonly change?: PropertyChange<TableProps>
}

export interface RowHeight {
	// Twips.
	readonly value: number
	readonly rule: "auto" | "exact" | "atLeast"
}

// w:trPr (§17.4.82).
export interface TableRowProps {
	readonly conditional?: ConditionalFlags
	readonly divId?: number
	readonly gridBefore?: number
	readonly gridAfter?: number
	readonly widthBefore?: WidthProps
	readonly widthAfter?: WidthProps
	readonly cantSplit?: boolean
	readonly height?: RowHeight
	readonly header?: boolean
	readonly cellSpacing?: WidthProps
	readonly alignment?: string
	readonly hidden?: boolean
	readonly revisions?: readonly RevisionMark[]
	readonly change?: PropertyChange<TableRowProps>
}

export interface CellBordersProps extends TableBordersProps {
	readonly tl2br?: BorderProps
	readonly tr2bl?: BorderProps
}

// w:tcPr (§17.4.70).
export interface TableCellProps {
	readonly conditional?: ConditionalFlags
	readonly width?: WidthProps
	readonly gridSpan?: number
	readonly hMerge?: "restart" | "continue"
	readonly vMerge?: "restart" | "continue"
	readonly borders?: CellBordersProps
	readonly shading?: ShadingProps
	readonly noWrap?: boolean
	readonly margins?: CellMarginsProps
	readonly textDirection?: string
	readonly fitText?: boolean
	readonly vAlign?: "top" | "center" | "bottom" | "both"
	readonly hideMark?: boolean
	// w:cellIns / w:cellDel / w:cellMerge.
	readonly revisions?: readonly RevisionMark[]
	readonly change?: PropertyChange<TableCellProps>
}

export type HeaderFooterType = "default" | "first" | "even"

export interface HeaderFooterRef {
	readonly type: HeaderFooterType
	readonly relationshipId: string
}

export interface NotePositionProps {
	// ST_FtnPos / ST_EdnPos: "pageBottom", "beneathText", "sectEnd", "docEnd".
	readonly position?: string
	readonly numFormat?: string
	readonly numStart?: number
	// "continuous", "eachSect", "eachPage".
	readonly numRestart?: string
}

export interface PageSize {
	readonly width?: number
	readonly height?: number
	readonly orient?: "portrait" | "landscape"
	readonly code?: number
}

export interface PageMargins {
	readonly top?: number
	readonly right?: number
	readonly bottom?: number
	readonly left?: number
	readonly header?: number
	readonly footer?: number
	readonly gutter?: number
}

export interface PageBorders {
	readonly zOrder?: "front" | "back"
	readonly display?: "allPages" | "firstPage" | "notFirstPage"
	readonly offsetFrom?: "page" | "text"
	readonly top?: BorderProps
	readonly left?: BorderProps
	readonly bottom?: BorderProps
	readonly right?: BorderProps
}

export interface LineNumbering {
	readonly countBy?: number
	readonly start?: number
	readonly distance?: number
	readonly restart?: "newPage" | "newSection" | "continuous"
}

export interface PageNumbering {
	readonly format?: string
	readonly start?: number
	readonly chapterStyle?: number
	readonly chapterSeparator?: string
}

export interface Column {
	readonly width?: number
	readonly space?: number
}

export interface ColumnsProps {
	readonly count?: number
	readonly space?: number
	readonly equalWidth?: boolean
	readonly separator?: boolean
	readonly columns?: readonly Column[]
}

export interface DocGrid {
	// "default", "lines", "linesAndChars", "snapToChars".
	readonly type?: string
	// Twips.
	readonly linePitch?: number
	// ST_DecimalNumber as stored (a packed character pitch delta).
	readonly charSpace?: number
}

// w:sectPr (§17.6.17).
export interface SectionProps {
	readonly headers?: readonly HeaderFooterRef[]
	readonly footers?: readonly HeaderFooterRef[]
	readonly footnotes?: NotePositionProps
	readonly endnotes?: NotePositionProps
	// ST_SectionMark: "nextPage", "continuous", "evenPage", "oddPage", "nextColumn".
	readonly type?: string
	readonly pageSize?: PageSize
	readonly pageMargins?: PageMargins
	readonly pageBorders?: PageBorders
	readonly lineNumbering?: LineNumbering
	readonly pageNumbering?: PageNumbering
	readonly columns?: ColumnsProps
	readonly formProtection?: boolean
	readonly vAlign?: "top" | "center" | "both" | "bottom"
	readonly noEndnote?: boolean
	readonly titlePage?: boolean
	readonly textDirection?: string
	readonly bidi?: boolean
	readonly rtlGutter?: boolean
	readonly docGrid?: DocGrid
	readonly printerSettings?: string
	readonly change?: PropertyChange<SectionProps>
}

// ---- Resolved ----------------------------------------------------------------------------------------

// Colours are six upper-case hex digits; undefined means "auto" (the renderer's automatic colour).
export interface ResolvedBorder {
	readonly style: string
	readonly size: number
	readonly space: number
	readonly color: string | undefined
	readonly shadow: boolean
	readonly frame: boolean
}

export interface ResolvedShading {
	readonly pattern: string
	readonly color: string | undefined
	readonly fill: string | undefined
}

export interface ResolvedUnderline {
	readonly style: string
	readonly color: string | undefined
}

export interface ResolvedFonts {
	readonly ascii: string
	readonly hAnsi: string
	readonly eastAsia: string
	readonly cs: string
	readonly hint: "default" | "eastAsia" | "cs"
	// Word draws East Asian characters with the ascii font here (eastAsia is its Times New Roman default and
	// ascii equals hAnsi): core/fonts eastAsiaUsesAscii.
	readonly eastAsiaUsesAscii: boolean
}

export interface ResolvedLanguage {
	readonly latin: string | undefined
	readonly eastAsia: string | undefined
	readonly bidi: string | undefined
}

// The effective formatting of a run. `id` is dense per resolver: equal sets are one object.
export interface ResolvedRunProps {
	readonly id: number
	readonly styleId: string | undefined
	readonly fonts: ResolvedFonts
	readonly size: number
	readonly sizeCs: number
	readonly bold: boolean
	readonly boldCs: boolean
	readonly italic: boolean
	readonly italicCs: boolean
	readonly caps: boolean
	readonly smallCaps: boolean
	readonly strike: boolean
	readonly doubleStrike: boolean
	readonly outline: boolean
	readonly shadow: boolean
	readonly emboss: boolean
	readonly imprint: boolean
	readonly vanish: boolean
	readonly specVanish: boolean
	readonly webHidden: boolean
	readonly noProof: boolean
	readonly snapToGrid: boolean
	readonly rtl: boolean
	readonly cs: boolean
	readonly oMath: boolean
	readonly color: string | undefined
	readonly highlight: string | undefined
	readonly underline: ResolvedUnderline | undefined
	readonly vertAlign: "baseline" | "superscript" | "subscript"
	readonly position: number
	readonly spacing: number
	readonly scale: number
	readonly kern: number
	readonly shading: ResolvedShading | undefined
	readonly border: ResolvedBorder | undefined
	readonly effect: string | undefined
	readonly emphasis: string | undefined
	readonly fitText: FitTextProps | undefined
	readonly lang: ResolvedLanguage
	readonly eastAsianLayout: EastAsianLayoutProps | undefined
}

export interface ResolvedTabStop {
	readonly alignment: "start" | "center" | "end" | "decimal" | "bar" | "num"
	readonly position: number
	readonly leader: string
}

export interface ResolvedSpacing {
	readonly before: number
	readonly beforeLines: number | undefined
	readonly beforeAutospacing: boolean
	readonly after: number
	readonly afterLines: number | undefined
	readonly afterAutospacing: boolean
	readonly line: number
	readonly lineRule: "auto" | "exact" | "atLeast"
}

// firstLine is signed: negative for a hanging indent.
export interface ResolvedIndent {
	readonly start: number
	readonly end: number
	readonly firstLine: number
	readonly startChars: number | undefined
	readonly endChars: number | undefined
	readonly firstLineChars: number | undefined
}

export interface ResolvedParagraphBorders {
	readonly top: ResolvedBorder | undefined
	readonly start: ResolvedBorder | undefined
	readonly bottom: ResolvedBorder | undefined
	readonly end: ResolvedBorder | undefined
	readonly between: ResolvedBorder | undefined
	readonly bar: ResolvedBorder | undefined
}

export type PhysicalAlignment = "left" | "center" | "right" | "justify" | "distribute"

export interface ResolvedNumbering {
	readonly numId: number
	readonly ilvl: number
}

export interface ResolvedParagraphProps {
	readonly id: number
	readonly styleId: string | undefined
	// ST_Jc as resolved ("start" for none) and its physical alignment for the paragraph direction.
	readonly alignment: string
	readonly physicalAlignment: PhysicalAlignment
	readonly indent: ResolvedIndent
	readonly spacing: ResolvedSpacing
	readonly contextualSpacing: boolean
	readonly keepNext: boolean
	readonly keepLines: boolean
	readonly widowControl: boolean
	readonly pageBreakBefore: boolean
	// Sorted by position, clears applied.
	readonly tabs: readonly ResolvedTabStop[]
	readonly numbering: ResolvedNumbering | undefined
	readonly outlineLevel: number | undefined
	readonly borders: ResolvedParagraphBorders | undefined
	readonly shading: ResolvedShading | undefined
	readonly frame: FrameProps | undefined
	readonly textDirection: string | undefined
	readonly textAlignment: string
	readonly bidi: boolean
	readonly snapToGrid: boolean
	readonly suppressAutoHyphens: boolean
	readonly suppressLineNumbers: boolean
	readonly suppressOverlap: boolean
	readonly mirrorIndents: boolean
	readonly kinsoku: boolean
	readonly wordWrap: boolean
	readonly overflowPunct: boolean
	readonly topLinePunct: boolean
	readonly autoSpaceDE: boolean
	readonly autoSpaceDN: boolean
	readonly adjustRightInd: boolean
}

export interface ResolvedTableBorders {
	readonly top: ResolvedBorder | undefined
	readonly start: ResolvedBorder | undefined
	readonly bottom: ResolvedBorder | undefined
	readonly end: ResolvedBorder | undefined
	readonly insideH: ResolvedBorder | undefined
	readonly insideV: ResolvedBorder | undefined
}

export interface ResolvedCellMargins {
	readonly top: WidthProps
	readonly start: WidthProps
	readonly bottom: WidthProps
	readonly end: WidthProps
}

export interface ResolvedTableProps {
	readonly id: number
	readonly styleId: string | undefined
	readonly width: WidthProps
	readonly alignment: string
	readonly physicalAlignment: "left" | "center" | "right"
	readonly indent: WidthProps
	readonly cellSpacing: WidthProps
	readonly borders: ResolvedTableBorders
	readonly shading: ResolvedShading | undefined
	readonly layout: "fixed" | "autofit"
	readonly cellMargins: ResolvedCellMargins
	readonly look: TableLook
	readonly bidiVisual: boolean
	readonly overlap: "never" | "overlap"
	readonly position: TablePositionProps | undefined
	readonly rowBandSize: number
	readonly colBandSize: number
	readonly caption: string | undefined
	readonly description: string | undefined
}

export interface ResolvedTableRowProps {
	readonly id: number
	readonly gridBefore: number
	readonly gridAfter: number
	readonly widthBefore: WidthProps | undefined
	readonly widthAfter: WidthProps | undefined
	readonly cantSplit: boolean
	readonly height: RowHeight | undefined
	readonly header: boolean
	readonly hidden: boolean
	readonly cellSpacing: WidthProps | undefined
	readonly alignment: string | undefined
}

export interface ResolvedCellBorders extends ResolvedTableBorders {
	readonly tl2br: ResolvedBorder | undefined
	readonly tr2bl: ResolvedBorder | undefined
}

export interface ResolvedTableCellProps {
	readonly id: number
	readonly width: WidthProps | undefined
	readonly gridSpan: number
	readonly hMerge: "restart" | "continue" | undefined
	readonly vMerge: "restart" | "continue" | undefined
	readonly borders: ResolvedCellBorders
	readonly shading: ResolvedShading | undefined
	readonly margins: ResolvedCellMargins
	readonly noWrap: boolean
	readonly textDirection: string | undefined
	readonly fitText: boolean
	readonly vAlign: "top" | "center" | "bottom" | "both"
	readonly hideMark: boolean
}

// A section with Word's defaults filled in (US Letter, 1" margins, one column).
export interface ResolvedSectionProps {
	readonly type: string
	readonly pageWidth: number
	readonly pageHeight: number
	readonly orient: "portrait" | "landscape"
	readonly margins: Required<PageMargins>
	readonly columns: ColumnsProps & {
		readonly count: number
		readonly space: number
		readonly equalWidth: boolean
		readonly separator: boolean
	}
	readonly titlePage: boolean
	readonly vAlign: "top" | "center" | "both" | "bottom"
	readonly pageNumbering: PageNumbering
	readonly lineNumbering: LineNumbering | undefined
	readonly pageBorders: PageBorders | undefined
	readonly docGrid: DocGrid
	readonly textDirection: string | undefined
	readonly bidi: boolean
	readonly rtlGutter: boolean
	readonly headers: readonly HeaderFooterRef[]
	readonly footers: readonly HeaderFooterRef[]
	readonly footnotes: NotePositionProps | undefined
	readonly endnotes: NotePositionProps | undefined
}
