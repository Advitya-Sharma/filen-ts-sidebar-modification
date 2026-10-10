// Namespace groups: each transitional URI and its Strict (ISO 29500 Strict) alias resolve to one group, so
// elements are compared as (group, local name) and never by prefix.

export const Ns = {
	None: 0,
	W: 1,
	R: 2,
	W14: 3,
	W15: 4,
	W16Cid: 5,
	W16Cex: 6,
	MC: 7,
	WP: 8,
	A: 9,
	Pic: 10,
	C: 11,
	Dgm: 12,
	Wps: 13,
	Wpg: 14,
	Wpc: 15,
	Wpi: 16,
	M: 17,
	V: 18,
	O: 19,
	W10: 20,
	Dsp: 21,
	Xml: 22
} as const

export type Ns = (typeof Ns)[keyof typeof Ns]

const GROUP_URIS: readonly (readonly [Ns, readonly string[]])[] = [
	[Ns.W, ["http://schemas.openxmlformats.org/wordprocessingml/2006/main", "http://purl.oclc.org/ooxml/wordprocessingml/main"]],
	[
		Ns.R,
		["http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships"]
	],
	[Ns.W14, ["http://schemas.microsoft.com/office/word/2010/wordml"]],
	[Ns.W15, ["http://schemas.microsoft.com/office/word/2012/wordml"]],
	[Ns.W16Cid, ["http://schemas.microsoft.com/office/word/2016/wordml/cid"]],
	[Ns.W16Cex, ["http://schemas.microsoft.com/office/word/2018/wordml/cex"]],
	[Ns.MC, ["http://schemas.openxmlformats.org/markup-compatibility/2006"]],
	[
		Ns.WP,
		[
			"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
			"http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing"
		]
	],
	[Ns.A, ["http://schemas.openxmlformats.org/drawingml/2006/main", "http://purl.oclc.org/ooxml/drawingml/main"]],
	[Ns.Pic, ["http://schemas.openxmlformats.org/drawingml/2006/picture", "http://purl.oclc.org/ooxml/drawingml/picture"]],
	[Ns.C, ["http://schemas.openxmlformats.org/drawingml/2006/chart", "http://purl.oclc.org/ooxml/drawingml/chart"]],
	[Ns.Dgm, ["http://schemas.openxmlformats.org/drawingml/2006/diagram", "http://purl.oclc.org/ooxml/drawingml/diagram"]],
	[Ns.Wps, ["http://schemas.microsoft.com/office/word/2010/wordprocessingShape"]],
	[Ns.Wpg, ["http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"]],
	[Ns.Wpc, ["http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"]],
	[Ns.Wpi, ["http://schemas.microsoft.com/office/word/2010/wordprocessingInk"]],
	[Ns.M, ["http://schemas.openxmlformats.org/officeDocument/2006/math", "http://purl.oclc.org/ooxml/officeDocument/math"]],
	[Ns.V, ["urn:schemas-microsoft-com:vml"]],
	[Ns.O, ["urn:schemas-microsoft-com:office:office"]],
	[Ns.W10, ["urn:schemas-microsoft-com:office:word"]],
	[Ns.Dsp, ["http://schemas.microsoft.com/office/drawing/2008/diagram"]],
	[Ns.Xml, ["http://www.w3.org/XML/1998/namespace"]]
]

const GROUP_BY_URI = new Map<string, Ns>()

for (const [group, uris] of GROUP_URIS) {
	for (const uri of uris) GROUP_BY_URI.set(uri, group)
}

export const URIS_OF_GROUP: readonly (readonly string[])[] = (() => {
	const out: string[][] = []

	for (const [group, uris] of GROUP_URIS) out[group] = [...uris]

	return out
})()

export function groupOf(uri: string | undefined): Ns {
	return uri === undefined ? Ns.None : (GROUP_BY_URI.get(uri) ?? Ns.None)
}

export const STRICT_WORDPROCESSINGML = "http://purl.oclc.org/ooxml/wordprocessingml/main"

// The graphicData kinds a drawing can carry, by a:graphicData@uri.
export const GRAPHIC_URIS: ReadonlyMap<string, GraphicKind> = new Map<string, GraphicKind>([
	["http://schemas.openxmlformats.org/drawingml/2006/picture", "picture"],
	["http://purl.oclc.org/ooxml/drawingml/picture", "picture"],
	["http://schemas.microsoft.com/office/word/2010/wordprocessingShape", "shape"],
	["http://schemas.microsoft.com/office/word/2010/wordprocessingGroup", "group"],
	["http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas", "canvas"],
	["http://schemas.openxmlformats.org/drawingml/2006/chart", "chart"],
	["http://purl.oclc.org/ooxml/drawingml/chart", "chart"],
	["http://schemas.microsoft.com/office/drawing/2014/chartex", "chartEx"],
	["http://schemas.openxmlformats.org/drawingml/2006/diagram", "smartArt"],
	["http://purl.oclc.org/ooxml/drawingml/diagram", "smartArt"],
	["http://schemas.microsoft.com/office/word/2010/wordprocessingInk", "ink"],
	["http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas", "lockedCanvas"]
])

export type GraphicKind = "picture" | "shape" | "group" | "canvas" | "chart" | "chartEx" | "smartArt" | "ink" | "lockedCanvas" | "unknown"

// mc:Choice branches are taken only when every required namespace is one the model reads; anything else
// (ink, chartex, 3D models) falls back to the producer's own fallback, usually a picture.
const UNDERSTOOD_CHOICE_URIS = new Set<string>([
	...(URIS_OF_GROUP[Ns.Wps] ?? []),
	...(URIS_OF_GROUP[Ns.Wpg] ?? []),
	...(URIS_OF_GROUP[Ns.W14] ?? []),
	...(URIS_OF_GROUP[Ns.W15] ?? []),
	...(URIS_OF_GROUP[Ns.M] ?? []),
	"http://schemas.microsoft.com/office/drawing/2010/main",
	"http://schemas.microsoft.com/office/word/2015/wordml/symex"
])

export function understandsChoice(uri: string): boolean {
	return UNDERSTOOD_CHOICE_URIS.has(uri)
}

// Element codes. Only the names the model branches on get one; everything else is El.Unknown.
export const El = {
	Unknown: 0,
	// w: structure
	Document: 1,
	Body: 2,
	Background: 3,
	Header: 4,
	Footer: 5,
	Footnotes: 6,
	Endnotes: 7,
	Footnote: 8,
	Endnote: 9,
	Comments: 10,
	Comment: 11,
	TxbxContent: 12,
	// w: blocks
	P: 13,
	PPr: 14,
	Tbl: 15,
	TblPr: 16,
	TblGrid: 17,
	Tr: 18,
	TrPr: 19,
	TblPrEx: 20,
	Tc: 21,
	TcPr: 22,
	SectPr: 23,
	AltChunk: 24,
	SubDoc: 25,
	// w: wrappers
	Sdt: 26,
	SdtPr: 27,
	SdtEndPr: 28,
	SdtContent: 29,
	CustomXml: 30,
	CustomXmlPr: 31,
	SmartTag: 32,
	SmartTagPr: 33,
	Hyperlink: 34,
	FldSimple: 35,
	Ins: 36,
	Del: 37,
	MoveFrom: 38,
	MoveTo: 39,
	Dir: 40,
	Bdo: 41,
	ProofErr: 42,
	// w: run content
	R: 43,
	RPr: 44,
	T: 45,
	DelText: 46,
	InstrText: 47,
	DelInstrText: 48,
	Tab: 49,
	Br: 50,
	Cr: 51,
	Sym: 52,
	SoftHyphen: 53,
	NoBreakHyphen: 54,
	Ptab: 55,
	LastRenderedPageBreak: 56,
	FldChar: 57,
	FootnoteReference: 58,
	EndnoteReference: 59,
	FootnoteRef: 60,
	EndnoteRef: 61,
	Separator: 62,
	ContinuationSeparator: 63,
	CommentReference: 64,
	AnnotationRef: 65,
	Drawing: 66,
	Pict: 67,
	Object: 68,
	Ruby: 69,
	RubyPr: 70,
	Rt: 71,
	RubyBase: 72,
	PgNum: 73,
	DayShort: 74,
	MonthShort: 75,
	YearShort: 76,
	DayLong: 77,
	MonthLong: 78,
	YearLong: 79,
	ContentPart: 80,
	// w: markers
	BookmarkStart: 81,
	BookmarkEnd: 82,
	CommentRangeStart: 83,
	CommentRangeEnd: 84,
	MoveFromRangeStart: 85,
	MoveFromRangeEnd: 86,
	MoveToRangeStart: 87,
	MoveToRangeEnd: 88,
	PermStart: 89,
	PermEnd: 90,
	CustomXmlRangeMarker: 91,
	// w: properties and their revisions
	PStyle: 92,
	RStyle: 93,
	TblStyle: 94,
	NumPr: 95,
	RPrChange: 96,
	PPrChange: 97,
	SectPrChange: 98,
	TblPrChange: 99,
	TrPrChange: 100,
	TcPrChange: 101,
	TblGridChange: 102,
	TblPrExChange: 103,
	NumberingChange: 104,
	CellIns: 105,
	CellDel: 106,
	CellMerge: 107,
	HeaderReference: 108,
	FooterReference: 109,
	// mc, m, w14
	AlternateContent: 120,
	Choice: 121,
	Fallback: 122,
	OMath: 123,
	OMathPara: 124,
	W14ContentPart: 125
} as const

export type El = (typeof El)[keyof typeof El]

const W_CODES: readonly (readonly [string, El])[] = [
	["document", El.Document],
	["body", El.Body],
	["background", El.Background],
	["hdr", El.Header],
	["ftr", El.Footer],
	["footnotes", El.Footnotes],
	["endnotes", El.Endnotes],
	["footnote", El.Footnote],
	["endnote", El.Endnote],
	["comments", El.Comments],
	["comment", El.Comment],
	["txbxContent", El.TxbxContent],
	["p", El.P],
	["pPr", El.PPr],
	["tbl", El.Tbl],
	["tblPr", El.TblPr],
	["tblGrid", El.TblGrid],
	["tr", El.Tr],
	["trPr", El.TrPr],
	["tblPrEx", El.TblPrEx],
	["tc", El.Tc],
	["tcPr", El.TcPr],
	["sectPr", El.SectPr],
	["altChunk", El.AltChunk],
	["subDoc", El.SubDoc],
	["sdt", El.Sdt],
	["sdtPr", El.SdtPr],
	["sdtEndPr", El.SdtEndPr],
	["sdtContent", El.SdtContent],
	["customXml", El.CustomXml],
	["customXmlPr", El.CustomXmlPr],
	["smartTag", El.SmartTag],
	["smartTagPr", El.SmartTagPr],
	["hyperlink", El.Hyperlink],
	["fldSimple", El.FldSimple],
	["ins", El.Ins],
	["del", El.Del],
	["moveFrom", El.MoveFrom],
	["moveTo", El.MoveTo],
	["dir", El.Dir],
	["bdo", El.Bdo],
	["proofErr", El.ProofErr],
	["r", El.R],
	["rPr", El.RPr],
	["t", El.T],
	["delText", El.DelText],
	["instrText", El.InstrText],
	["delInstrText", El.DelInstrText],
	["tab", El.Tab],
	["br", El.Br],
	["cr", El.Cr],
	["sym", El.Sym],
	["softHyphen", El.SoftHyphen],
	["noBreakHyphen", El.NoBreakHyphen],
	["ptab", El.Ptab],
	["lastRenderedPageBreak", El.LastRenderedPageBreak],
	["fldChar", El.FldChar],
	["footnoteReference", El.FootnoteReference],
	["endnoteReference", El.EndnoteReference],
	["footnoteRef", El.FootnoteRef],
	["endnoteRef", El.EndnoteRef],
	["separator", El.Separator],
	["continuationSeparator", El.ContinuationSeparator],
	["commentReference", El.CommentReference],
	["annotationRef", El.AnnotationRef],
	["drawing", El.Drawing],
	["pict", El.Pict],
	["object", El.Object],
	["ruby", El.Ruby],
	["rubyPr", El.RubyPr],
	["rt", El.Rt],
	["rubyBase", El.RubyBase],
	["pgNum", El.PgNum],
	["dayShort", El.DayShort],
	["monthShort", El.MonthShort],
	["yearShort", El.YearShort],
	["dayLong", El.DayLong],
	["monthLong", El.MonthLong],
	["yearLong", El.YearLong],
	["contentPart", El.ContentPart],
	["bookmarkStart", El.BookmarkStart],
	["bookmarkEnd", El.BookmarkEnd],
	["commentRangeStart", El.CommentRangeStart],
	["commentRangeEnd", El.CommentRangeEnd],
	["moveFromRangeStart", El.MoveFromRangeStart],
	["moveFromRangeEnd", El.MoveFromRangeEnd],
	["moveToRangeStart", El.MoveToRangeStart],
	["moveToRangeEnd", El.MoveToRangeEnd],
	["permStart", El.PermStart],
	["permEnd", El.PermEnd],
	["customXmlInsRangeStart", El.CustomXmlRangeMarker],
	["customXmlInsRangeEnd", El.CustomXmlRangeMarker],
	["customXmlDelRangeStart", El.CustomXmlRangeMarker],
	["customXmlDelRangeEnd", El.CustomXmlRangeMarker],
	["customXmlMoveFromRangeStart", El.CustomXmlRangeMarker],
	["customXmlMoveFromRangeEnd", El.CustomXmlRangeMarker],
	["customXmlMoveToRangeStart", El.CustomXmlRangeMarker],
	["customXmlMoveToRangeEnd", El.CustomXmlRangeMarker],
	["pStyle", El.PStyle],
	["rStyle", El.RStyle],
	["tblStyle", El.TblStyle],
	["numPr", El.NumPr],
	["rPrChange", El.RPrChange],
	["pPrChange", El.PPrChange],
	["sectPrChange", El.SectPrChange],
	["tblPrChange", El.TblPrChange],
	["trPrChange", El.TrPrChange],
	["tcPrChange", El.TcPrChange],
	["tblGridChange", El.TblGridChange],
	["tblPrExChange", El.TblPrExChange],
	["numberingChange", El.NumberingChange],
	["cellIns", El.CellIns],
	["cellDel", El.CellDel],
	["cellMerge", El.CellMerge],
	["headerReference", El.HeaderReference],
	["footerReference", El.FooterReference]
]

export const CODES_BY_GROUP: readonly ReadonlyMap<string, El>[] = (() => {
	const out: Map<string, El>[] = []

	for (let group = 0; group <= Ns.Xml; group++) out.push(new Map())

	out[Ns.W] = new Map(W_CODES)
	out[Ns.MC] = new Map<string, El>([
		["AlternateContent", El.AlternateContent],
		["Choice", El.Choice],
		["Fallback", El.Fallback]
	])
	out[Ns.M] = new Map<string, El>([
		["oMath", El.OMath],
		["oMathPara", El.OMathPara]
	])
	out[Ns.W14] = new Map<string, El>([["contentPart", El.W14ContentPart]])

	return out
})()
