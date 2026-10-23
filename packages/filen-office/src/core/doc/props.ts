import { le16, le32, les16, les32 } from "./bytes"
import { forEachPrl, SGC_CHARACTER, SGC_PARAGRAPH, SGC_SECTION, SGC_TABLE, sprmSgc } from "./sprms"

// Typed property sets for [MS-DOC] §2.6 sprms. Every sprm the spec lists is recognised (sprms.ts); the ones below
// change a typed property. The rest stay in the run's or paragraph's raw grpprl, never dropped.

export type DocColor = number | "auto"

export interface DocBorder {
	// BrcType (§2.9.17); 0 = none.
	readonly type: number
	// Eighths of a point.
	readonly width: number
	readonly color: DocColor
	// Points.
	readonly space: number
	readonly shadow: boolean
	readonly frame: boolean
}

export interface DocShading {
	readonly fore: DocColor
	readonly back: DocColor
	// Ipat (§2.9.124).
	readonly pattern: number
}

// Ico (§2.9.119) colours.
const ICO: readonly number[] = [
	-1, 0x000000, 0x0000ff, 0x00ffff, 0x00ff00, 0xff00ff, 0xff0000, 0xffff00, 0xffffff, 0x000080, 0x008080, 0x008000, 0x800080, 0x800000,
	0x808000, 0x808080, 0xc0c0c0
]

export function icoColor(ico: number): DocColor {
	const rgb = ICO[ico] ?? -1

	return rgb < 0 ? "auto" : rgb
}

// COLORREF (§2.9.39): red, green, blue, fAuto.
export function colorRef(data: Uint8Array, at: number): DocColor {
	if ((data[at + 3] ?? 0) === 0xff) return "auto"

	return ((data[at] ?? 0) << 16) | ((data[at + 1] ?? 0) << 8) | (data[at + 2] ?? 0)
}

// Brc80 (§2.9.17 Brc80): dptLineWidth, brcType, ico, dptSpace:5 fShadow fFrame.
export function brc80(data: Uint8Array, at: number): DocBorder | undefined {
	// Brc80MayBeNil: 0xFFFFFFFF is "no border specified".
	if (le32(data, at) === 0xffffffff) return undefined

	const flags = data[at + 3] ?? 0

	return {
		width: data[at] ?? 0,
		type: data[at + 1] ?? 0,
		color: icoColor(data[at + 2] ?? 0),
		space: flags & 0x1f,
		shadow: (flags & 0x20) !== 0,
		frame: (flags & 0x40) !== 0
	}
}

// Brc (§2.9.16): cv, dptLineWidth, brcType, dptSpace:5 fShadow fFrame.
export function brc(data: Uint8Array, at: number): DocBorder | undefined {
	if (le32(data, at + 4) === 0xffffffff) return undefined

	const flags = le16(data, at + 6)

	return {
		color: colorRef(data, at),
		width: data[at + 4] ?? 0,
		type: data[at + 5] ?? 0,
		space: flags & 0x1f,
		shadow: (flags & 0x20) !== 0,
		frame: (flags & 0x40) !== 0
	}
}

// Shd80 (§2.9.249): icoFore:5 icoBack:5 ipat:6.
export function shd80(value: number): DocShading {
	return { fore: icoColor(value & 0x1f), back: icoColor((value >> 5) & 0x1f), pattern: value >> 10 }
}

// Shd (§2.9.248): cvFore, cvBack, ipat.
export function shd(data: Uint8Array, at: number): DocShading {
	return { fore: colorRef(data, at), back: colorRef(data, at + 4), pattern: le16(data, at + 8) }
}

// DTTM (§2.9.65) as an ISO 8601 local time, or undefined when zero.
export function dttm(value: number): string | undefined {
	if (value === 0) return undefined

	const minute = value & 0x3f
	const hour = (value >>> 6) & 0x1f
	const day = (value >>> 11) & 0x1f
	const month = (value >>> 16) & 0xf
	const year = 1900 + ((value >>> 20) & 0x1ff)
	const two = (n: number): string => String(n).padStart(2, "0")

	if (month < 1 || month > 12 || day < 1) return undefined

	return `${year}-${two(month)}-${two(day)}T${two(hour)}:${two(minute)}:00`
}

export interface DocRevision {
	// Index into the revision authors (SttbfRMark).
	readonly author: number
	readonly date: string | undefined
}

export interface DocChar {
	// Character style istd (sprmCIstd); 10 is Default Paragraph Font.
	istd: number
	bold: boolean
	italic: boolean
	strike: boolean
	doubleStrike: boolean
	outline: boolean
	shadow: boolean
	emboss: boolean
	imprint: boolean
	smallCaps: boolean
	caps: boolean
	vanish: boolean
	specVanish: boolean
	webHidden: boolean
	fieldVanish: boolean
	boldBi: boolean
	italicBi: boolean
	rtl: boolean
	complexScripts: boolean
	noProof: boolean
	// Kul (§2.9.135); 0 = none.
	underline: number
	underlineColor: DocColor
	color: DocColor
	highlight: DocColor | undefined
	// Half-points.
	size: number
	sizeBi: number
	// Indexes into the font table.
	fontAscii: number
	fontEastAsia: number
	fontOther: number
	fontBi: number
	lang: number
	langEastAsia: number
	langBi: number
	// 0 normal, 1 superscript, 2 subscript.
	vertAlign: number
	// Half-points, positive raises.
	position: number
	// Twips.
	spacing: number
	kern: number
	// Percent.
	scale: number
	emphasis: number
	effect: number
	shading: DocShading | undefined
	border: DocBorder | undefined
	symbol: { readonly font: number; readonly char: number } | undefined
	// fSpec: the character is a special character (picture, field mark, footnote reference, ...).
	special: boolean
	// fObj / fOle2 / fData (embedded object, OLE 2 object, binary form field data).
	object: boolean
	ole2: boolean
	data: boolean
	// sprmCPicLocation: offset in the Data stream (or the ObjectPool storage name for OLE objects).
	picLocation: number | undefined
	inserted: DocRevision | undefined
	deleted: DocRevision | undefined
	formatChange: DocRevision | undefined
	// sprmCIbstRMark / sprmCDttmRMark seen before the insertion flag, kept until the flag arrives.
	insertAuthor: number
	insertDate: string | undefined
	deleteAuthor: number
	deleteDate: string | undefined
}

export function defaultChar(): DocChar {
	return {
		istd: 10,
		bold: false,
		italic: false,
		strike: false,
		doubleStrike: false,
		outline: false,
		shadow: false,
		emboss: false,
		imprint: false,
		smallCaps: false,
		caps: false,
		vanish: false,
		specVanish: false,
		webHidden: false,
		fieldVanish: false,
		boldBi: false,
		italicBi: false,
		rtl: false,
		complexScripts: false,
		noProof: false,
		underline: 0,
		underlineColor: "auto",
		color: "auto",
		highlight: undefined,
		size: 20,
		sizeBi: 20,
		fontAscii: 0,
		fontEastAsia: 0,
		fontOther: 0,
		fontBi: 0,
		lang: 0x0400,
		langEastAsia: 0x0400,
		langBi: 0x0400,
		vertAlign: 0,
		position: 0,
		spacing: 0,
		kern: 0,
		scale: 100,
		emphasis: 0,
		effect: 0,
		shading: undefined,
		border: undefined,
		symbol: undefined,
		special: false,
		object: false,
		ole2: false,
		data: false,
		picLocation: undefined,
		inserted: undefined,
		deleted: undefined,
		formatChange: undefined,
		insertAuthor: 0,
		insertDate: undefined,
		deleteAuthor: 0,
		deleteDate: undefined
	}
}

type BooleanKeys<T> = { [K in keyof T]: T[K] extends boolean ? K : never }[keyof T]

// Toggle sprms (spra 0, sgc 2) by property.
const CHAR_TOGGLES: ReadonlyMap<number, BooleanKeys<DocChar>> = new Map<number, BooleanKeys<DocChar>>([
	[0x0802, "fieldVanish"],
	[0x0811, "webHidden"],
	[0x0818, "specVanish"],
	[0x0835, "bold"],
	[0x0836, "italic"],
	[0x0837, "strike"],
	[0x0838, "outline"],
	[0x0839, "shadow"],
	[0x083a, "smallCaps"],
	[0x083b, "caps"],
	[0x083c, "vanish"],
	[0x0854, "imprint"],
	[0x0855, "special"],
	[0x0856, "object"],
	[0x0858, "emboss"],
	[0x085a, "rtl"],
	[0x085c, "boldBi"],
	[0x085d, "italicBi"],
	[0x0875, "noProof"],
	[0x0882, "complexScripts"],
	[0x080a, "ole2"],
	[0x0806, "data"]
])

// Characters properties that sprmCPlain keeps (§2.6.1 sprmCPlain).
const PLAIN_KEEPS: readonly (keyof DocChar)[] = [
	"inserted",
	"deleted",
	"formatChange",
	"insertAuthor",
	"insertDate",
	"deleteAuthor",
	"deleteDate",
	"fieldVanish",
	"special",
	"object",
	"ole2",
	"data",
	"picLocation",
	"rtl",
	"complexScripts"
]

export interface CharContext {
	// The properties the text has from its styles alone (for ToggleOperand 0x80/0x81 and sprmCPlain).
	readonly style: DocChar
	// The paragraph style's character properties (sprmCPlain resets to these); `style` when absent.
	readonly paragraphStyle?: DocChar
	// Told about each sprm of this kind that no typed property takes.
	readonly untyped?: (sprm: number) => void
}

function toggle(operand: number, styleValue: boolean): boolean {
	if (operand === 0x80) return styleValue
	if (operand === 0x81) return !styleValue

	return operand !== 0
}

// Applies the character sprms of data[start, end) to `target`.
export function applyChar(target: DocChar, data: Uint8Array, start: number, end: number, context: CharContext): void {
	forEachPrl(data, start, end, (sprm, at, length) => {
		if (sprmSgc(sprm) !== SGC_CHARACTER) return

		const toggled = CHAR_TOGGLES.get(sprm)

		if (toggled !== undefined) {
			target[toggled] = toggle(data[at] ?? 0, context.style[toggled])
			return
		}

		const b = data[at] ?? 0
		const w = le16(data, at)

		switch (sprm) {
			case 0x0800:
				target.deleted = b === 0 ? undefined : { author: target.deleteAuthor, date: target.deleteDate }
				break
			case 0x0801:
				target.inserted = b === 0 ? undefined : { author: target.insertAuthor, date: target.insertDate }
				break
			case 0x4804:
				target.insertAuthor = les16(data, at)
				if (target.inserted !== undefined) target.inserted = { ...target.inserted, author: target.insertAuthor }
				break
			case 0x6805:
				target.insertDate = dttm(le32(data, at))
				if (target.inserted !== undefined) target.inserted = { ...target.inserted, date: target.insertDate }
				break
			case 0x4863:
				target.deleteAuthor = les16(data, at)
				if (target.deleted !== undefined) target.deleted = { ...target.deleted, author: target.deleteAuthor }
				break
			case 0x6864:
				target.deleteDate = dttm(le32(data, at))
				if (target.deleted !== undefined) target.deleted = { ...target.deleted, date: target.deleteDate }
				break
			case 0xca89:
			case 0xca57:
				// PropRMark (§2.9.217) after the size byte: fPropRMark, ibstshort, dttm.
				if (length >= 8) {
					target.formatChange =
						(data[at + 1] ?? 0) === 0 ? undefined : { author: les16(data, at + 2), date: dttm(le32(data, at + 4)) }
				}
				break
			case 0x6a03:
				target.picLocation = le32(data, at)
				break
			case 0x6a09:
				target.symbol = { font: w, char: le16(data, at + 2) }
				break
			case 0x2a0c:
				target.highlight = b === 0 ? undefined : icoColor(b)
				break
			case 0x4a30:
				// The run's base already holds this character style's properties (StyleResolver.charBase).
				target.istd = w
				break
			case 0x2a33:
				plain(target, context.paragraphStyle ?? context.style)
				break
			case 0x2a34:
				target.emphasis = b
				break
			case 0x2a3e:
				target.underline = b
				break
			case 0x8840:
				target.spacing = les16(data, at)
				break
			case 0x2a42:
				target.color = icoColor(b)
				break
			case 0x6870:
				target.color = colorRef(data, at)
				break
			case 0x6877:
				target.underlineColor = colorRef(data, at)
				break
			case 0x4a43:
				target.size = w
				break
			case 0x4a61:
				target.sizeBi = w
				break
			case 0x4845:
				target.position = les16(data, at)
				break
			case 0x2a48:
				target.vertAlign = b
				break
			case 0x484b:
				target.kern = w
				break
			case 0x4a4f:
				target.fontAscii = w
				break
			case 0x4a50:
				target.fontEastAsia = w
				break
			case 0x4a51:
				target.fontOther = w
				break
			case 0x4a5e:
				target.fontBi = w
				break
			case 0x4852:
				target.scale = w
				break
			case 0x2a53:
				target.doubleStrike = toggle(b, context.style.doubleStrike)
				break
			case 0x2859:
				target.effect = b
				break
			case 0x486d:
			case 0x4873:
				target.lang = w
				break
			case 0x486e:
			case 0x4874:
				target.langEastAsia = w
				break
			case 0x485f:
				target.langBi = w
				break
			case 0x4866:
				target.shading = shd80(w)
				break
			case 0xca71:
				if (length >= 11) target.shading = shd(data, at + 1)
				break
			case 0x6865:
				target.border = brc80(data, at)
				break
			case 0xca72:
				if (length >= 9) target.border = brc(data, at + 1)
				break
			default:
				context.untyped?.(sprm)
		}
	})
}

function plain(target: DocChar, style: DocChar): void {
	const kept = PLAIN_KEEPS.map(key => [key, target[key]] as const)

	Object.assign(target, style)

	for (const [key, value] of kept) (target as unknown as Record<string, unknown>)[key] = value
}

export interface DocTab {
	// Twips.
	readonly position: number
	// TBD (§2.9.311) jc: 0 left, 1 center, 2 right, 3 decimal, 4 bar, 6 list.
	readonly align: number
	// tlc: 0 none, 1 dot, 2 hyphen, 3 line, 4 heavy, 5 middle dot.
	readonly leader: number
}

export interface DocFrame {
	// PositionCodeOperand (§2.9.205): pcVert:2 pcHorz:2.
	readonly anchor: number
	readonly x: number
	readonly y: number
	readonly width: number
	readonly height: number
	readonly wrap: number
	readonly xFromText: number
	readonly yFromText: number
}

export interface DocPara {
	istd: number
	// Logical justification (sprmPJc values; sprmPJc80 maps 3–5 to 3).
	jc: number
	keep: boolean
	keepNext: boolean
	pageBreakBefore: boolean
	widowControl: boolean
	contextualSpacing: boolean
	suppressLineNumbers: boolean
	suppressAutoHyphens: boolean
	rtl: boolean
	mirrorIndents: boolean
	// Twips.
	left: number
	right: number
	firstLine: number
	// Hundredths of a character, when given in character units.
	leftChars: number | undefined
	rightChars: number | undefined
	firstLineChars: number | undefined
	before: number
	after: number
	beforeAuto: boolean
	afterAuto: boolean
	// LSPD: dyaLine (negative = exact) and whether it is a multiple of 240.
	line: number
	lineMultiple: boolean
	ilvl: number
	ilfo: number
	outlineLevel: number
	inTable: boolean
	ttp: boolean
	itap: number
	innerCell: boolean
	innerTtp: boolean
	shading: DocShading | undefined
	borderTop: DocBorder | undefined
	borderLeft: DocBorder | undefined
	borderBottom: DocBorder | undefined
	borderRight: DocBorder | undefined
	borderBetween: DocBorder | undefined
	borderBar: DocBorder | undefined
	tabs: DocTab[]
	frame: DocFrame | undefined
	// DCS (§2.9.53): fdct:3 lines:5.
	dropCap: number
	// sprmPFNumRMIns: the list number is revision-marked as inserted.
	numberInserted: boolean
	formatChange: DocRevision | undefined
	// sprmPHugePapx: offset of further paragraph sprms in the Data stream.
	hugePapx: number | undefined
	// sprmPTableProps / sprmPTIstdInfo seen (table formatting in a paragraph's style).
	tableStyleHint: boolean
}

export function defaultPara(): DocPara {
	return {
		istd: 0,
		jc: 0,
		keep: false,
		keepNext: false,
		pageBreakBefore: false,
		widowControl: true,
		contextualSpacing: false,
		suppressLineNumbers: false,
		suppressAutoHyphens: false,
		rtl: false,
		mirrorIndents: false,
		left: 0,
		right: 0,
		firstLine: 0,
		leftChars: undefined,
		rightChars: undefined,
		firstLineChars: undefined,
		before: 0,
		after: 0,
		beforeAuto: false,
		afterAuto: false,
		line: 240,
		lineMultiple: true,
		ilvl: 0,
		ilfo: 0,
		outlineLevel: 9,
		inTable: false,
		ttp: false,
		itap: 0,
		innerCell: false,
		innerTtp: false,
		shading: undefined,
		borderTop: undefined,
		borderLeft: undefined,
		borderBottom: undefined,
		borderRight: undefined,
		borderBetween: undefined,
		borderBar: undefined,
		tabs: [],
		frame: undefined,
		dropCap: 0,
		numberInserted: false,
		formatChange: undefined,
		hugePapx: undefined,
		tableStyleHint: false
	}
}

const MAX_TABS = 64

function changeTabs(target: DocPara, data: Uint8Array, at: number, end: number): void {
	// PChgTabsDelClose (cTabs, rgdxaDel[], rgdxaClose[]) then PChgTabsAdd (cTabs, rgdxaAdd[], rgtbdAdd[]).
	const deletes = data[at] ?? 0
	const deleted: number[] = []

	for (let k = 0; k < deletes && at + 1 + 2 * k + 2 <= end; k++) deleted.push(les16(data, at + 1 + 2 * k))

	addTabs(target, data, at + 1 + 4 * deletes, end, deleted)
}

// PChgTabsPapxOperand (§2.9.183): PChgTabsDel (cTabs, rgdxaDel[]) then PChgTabsAdd.
function changeTabsPapx(target: DocPara, data: Uint8Array, at: number, end: number): void {
	const deletes = data[at] ?? 0
	const deleted: number[] = []

	for (let k = 0; k < deletes && at + 1 + 2 * k + 2 <= end; k++) deleted.push(les16(data, at + 1 + 2 * k))

	addTabs(target, data, at + 1 + 2 * deletes, end, deleted)
}

function addTabs(target: DocPara, data: Uint8Array, at: number, end: number, deleted: readonly number[]): void {
	const adds = at < end ? (data[at] ?? 0) : 0
	const tabs = target.tabs.filter(tab => !deleted.includes(tab.position))

	for (let k = 0; k < adds; k++) {
		const positionAt = at + 1 + 2 * k
		const tbdAt = at + 1 + 2 * adds + k

		if (tbdAt >= end) break

		const position = les16(data, positionAt)
		const tbd = data[tbdAt] ?? 0
		const existing = tabs.findIndex(tab => tab.position === position)
		const tab: DocTab = { position, align: tbd & 7, leader: (tbd >> 3) & 7 }

		if (existing >= 0) tabs[existing] = tab
		else tabs.push(tab)
	}

	tabs.sort((a, b) => a.position - b.position)
	target.tabs = tabs.slice(0, MAX_TABS)
}

function frame(target: DocPara): DocFrame {
	return target.frame ?? { anchor: 0, x: 0, y: 0, width: 0, height: 0, wrap: 0, xFromText: 0, yFromText: 0 }
}

// Applies the paragraph sprms of data[start, end) to `target`. Table sprms in a PAPX (row marks carry the TAP)
// are applied separately by applyTable.
export function applyPara(target: DocPara, data: Uint8Array, start: number, end: number, untyped?: (sprm: number) => void): void {
	forEachPrl(data, start, end, (sprm, at, length) => {
		if (sprmSgc(sprm) !== SGC_PARAGRAPH) return

		const b = data[at] ?? 0
		const w = le16(data, at)
		const s = les16(data, at)

		switch (sprm) {
			case 0x4600:
				target.istd = w
				break
			case 0x2602:
				// sprmPIncLvl: heading styles 1–9 shift, others shift the outline level.
				if (target.istd >= 1 && target.istd <= 9) target.istd = Math.min(9, Math.max(1, target.istd + ((b << 24) >> 24)))
				else target.outlineLevel = Math.min(9, Math.max(0, target.outlineLevel + ((b << 24) >> 24)))
				break
			case 0x2403:
				target.jc = b > 3 ? 3 : b
				break
			case 0x2461:
				target.jc = b
				break
			case 0x2405:
				target.keep = b !== 0
				break
			case 0x2406:
				target.keepNext = b !== 0
				break
			case 0x2407:
				target.pageBreakBefore = b !== 0
				break
			case 0x260a:
				target.ilvl = Math.min(8, b)
				break
			case 0x460b:
				target.ilfo = s
				break
			case 0x240c:
				target.suppressLineNumbers = b !== 0
				break
			case 0xc60d:
				changeTabsPapx(target, data, at + 1, at + length)
				break
			case 0xc615:
				changeTabs(target, data, at + 1, at + length)
				break
			case 0x840e:
			case 0x845d:
				target.right = s
				target.rightChars = undefined
				break
			case 0x840f:
			case 0x845e:
				target.left = s
				target.leftChars = undefined
				break
			case 0x4610:
			case 0x465f:
				target.left = Math.max(0, target.left + s)
				break
			case 0x8411:
			case 0x8460:
				target.firstLine = s
				target.firstLineChars = undefined
				break
			case 0x4455:
				target.rightChars = s
				break
			case 0x4456:
				target.leftChars = s
				break
			case 0x4457:
				target.firstLineChars = s
				break
			case 0x6412:
				target.line = s
				target.lineMultiple = les16(data, at + 2) === 1
				break
			case 0xa413:
				target.before = w
				break
			case 0xa414:
				target.after = w
				break
			case 0x245b:
				target.beforeAuto = b !== 0
				break
			case 0x245c:
				target.afterAuto = b !== 0
				break
			case 0x2416:
				target.inTable = b !== 0
				break
			case 0x2417:
				target.ttp = b !== 0
				break
			case 0x6649:
				target.itap = Math.max(0, Math.min(64, les32(data, at)))
				break
			case 0x664a:
				target.itap = Math.max(0, Math.min(64, target.itap + les32(data, at)))
				break
			case 0x244b:
				target.innerCell = b !== 0
				break
			case 0x244c:
				target.innerTtp = b !== 0
				break
			case 0x2640:
				target.outlineLevel = Math.min(9, b)
				break
			case 0x2441:
				target.rtl = b !== 0
				break
			case 0x2431:
				target.widowControl = b !== 0
				break
			case 0x246d:
				target.contextualSpacing = b !== 0
				break
			case 0x242a:
				target.suppressAutoHyphens = b !== 0
				break
			case 0x2470:
				target.mirrorIndents = b !== 0
				break
			case 0x442d:
				target.shading = shd80(w)
				break
			case 0xc64d:
				if (length >= 11) target.shading = shd(data, at + 1)
				break
			case 0x6424:
				target.borderTop = brc80(data, at)
				break
			case 0x6425:
				target.borderLeft = brc80(data, at)
				break
			case 0x6426:
				target.borderBottom = brc80(data, at)
				break
			case 0x6427:
				target.borderRight = brc80(data, at)
				break
			case 0x6428:
				target.borderBetween = brc80(data, at)
				break
			case 0x6629:
				target.borderBar = brc80(data, at)
				break
			case 0xc64e:
				if (length >= 9) target.borderTop = brc(data, at + 1)
				break
			case 0xc64f:
				if (length >= 9) target.borderLeft = brc(data, at + 1)
				break
			case 0xc650:
				if (length >= 9) target.borderBottom = brc(data, at + 1)
				break
			case 0xc651:
				if (length >= 9) target.borderRight = brc(data, at + 1)
				break
			case 0xc652:
				if (length >= 9) target.borderBetween = brc(data, at + 1)
				break
			case 0xc653:
				if (length >= 9) target.borderBar = brc(data, at + 1)
				break
			case 0x261b:
				target.frame = { ...frame(target), anchor: b }
				break
			case 0x8418:
				target.frame = { ...frame(target), x: s }
				break
			case 0x8419:
				target.frame = { ...frame(target), y: s }
				break
			case 0x841a:
				target.frame = { ...frame(target), width: w }
				break
			case 0x442b:
				target.frame = { ...frame(target), height: w }
				break
			case 0x2423:
				target.frame = { ...frame(target), wrap: b }
				break
			case 0x842f:
				target.frame = { ...frame(target), xFromText: s }
				break
			case 0x842e:
				target.frame = { ...frame(target), yFromText: s }
				break
			case 0x442c:
				target.dropCap = w
				break
			case 0x2443:
				target.numberInserted = b !== 0
				break
			case 0xc66f:
				if (length >= 8) {
					target.formatChange =
						(data[at + 1] ?? 0) === 0 ? undefined : { author: les16(data, at + 2), date: dttm(le32(data, at + 4)) }
				}
				break
			case 0x6646:
				target.hugePapx = le32(data, at)
				break
			case 0x646b:
			case 0xc66c:
				target.tableStyleHint = true
				break
			default:
				untyped?.(sprm)
		}
	})
}

// TC80 / cell properties (§2.9.313, TCGRF §2.9.314).
export interface DocCell {
	// Twips from the row's left edge to the cell's right edge, from rgdxaCenter.
	readonly right: number
	readonly left: number
	// horzMerge: 0 none, 1 first of a merged range, 2 merged into the previous cell (or 3).
	horzMerge: number
	// vertMerge: 0 none, 2 continue, 3 restart (VerticalMergeFlag).
	vertMerge: number
	// 0 top, 1 center, 2 bottom.
	vertAlign: number
	textFlow: number
	widthUnit: number
	width: number
	noWrap: boolean
	fitText: boolean
	hideMark: boolean
	borderTop: DocBorder | undefined
	borderLeft: DocBorder | undefined
	borderBottom: DocBorder | undefined
	borderRight: DocBorder | undefined
	shading: DocShading | undefined
}

export interface DocWidth {
	// Fts (§2.9.103): 0 nil, 1 auto, 2 fifty-percent units, 3 twips.
	readonly unit: number
	readonly value: number
}

export interface DocRow {
	// Table justification: 0 left, 1 center, 2 right.
	jc: number
	left: number
	gapHalf: number
	// Negative = exact height, positive = at least.
	height: number
	cantSplit: boolean
	header: boolean
	rtl: boolean
	autofit: boolean
	istd: number | undefined
	width: DocWidth | undefined
	indent: DocWidth | undefined
	widthBefore: DocWidth | undefined
	widthAfter: DocWidth | undefined
	borderTop: DocBorder | undefined
	borderLeft: DocBorder | undefined
	borderBottom: DocBorder | undefined
	borderRight: DocBorder | undefined
	borderInsideH: DocBorder | undefined
	borderInsideV: DocBorder | undefined
	// CSSA defaults, twips: top, left, bottom, right.
	padding: [number | undefined, number | undefined, number | undefined, number | undefined]
	spacing: number | undefined
	cells: DocCell[]
	// Floating table position (sprmTPc ...), when any is set.
	positioned: boolean
	formatChange: DocRevision | undefined
}

export function defaultRow(): DocRow {
	return {
		jc: 0,
		left: 0,
		gapHalf: 0,
		height: 0,
		cantSplit: false,
		header: false,
		rtl: false,
		autofit: true,
		istd: undefined,
		width: undefined,
		indent: undefined,
		widthBefore: undefined,
		widthAfter: undefined,
		borderTop: undefined,
		borderLeft: undefined,
		borderBottom: undefined,
		borderRight: undefined,
		borderInsideH: undefined,
		borderInsideV: undefined,
		padding: [undefined, undefined, undefined, undefined],
		spacing: undefined,
		cells: [],
		positioned: false,
		formatChange: undefined
	}
}

const MAX_CELLS = 64

function emptyCell(left: number, right: number): DocCell {
	return {
		left,
		right,
		horzMerge: 0,
		vertMerge: 0,
		vertAlign: 0,
		textFlow: 0,
		widthUnit: 0,
		width: 0,
		noWrap: false,
		fitText: false,
		hideMark: false,
		borderTop: undefined,
		borderLeft: undefined,
		borderBottom: undefined,
		borderRight: undefined,
		shading: undefined
	}
}

function tcgrf(cell: DocCell, flags: number): void {
	cell.horzMerge = flags & 3
	cell.textFlow = (flags >> 2) & 7
	cell.vertMerge = (flags >> 5) & 3
	cell.vertAlign = (flags >> 7) & 3
	cell.widthUnit = (flags >> 9) & 7
	cell.fitText = (flags & 0x1000) !== 0
	cell.noWrap = (flags & 0x2000) !== 0
	cell.hideMark = (flags & 0x4000) !== 0
}

function width(data: Uint8Array, at: number): DocWidth {
	return { unit: data[at] ?? 0, value: les16(data, at + 1) }
}

function forCells(row: DocRow, first: number, lim: number, visit: (cell: DocCell) => void): void {
	for (let k = Math.max(0, first); k < Math.min(lim, row.cells.length); k++) {
		const cell = row.cells[k]

		if (cell !== undefined) visit(cell)
	}
}

function defTable(row: DocRow, data: Uint8Array, at: number, end: number): void {
	const columns = Math.min(MAX_CELLS - 1, data[at + 2] ?? 0)
	const centersAt = at + 3
	const tcAt = centersAt + 2 * (columns + 1)
	const cells: DocCell[] = []

	if (tcAt > end) return

	for (let k = 0; k < columns; k++) {
		const cell = emptyCell(les16(data, centersAt + 2 * k), les16(data, centersAt + 2 * (k + 1)))
		const tc = tcAt + 20 * k

		if (tc + 20 <= end) {
			tcgrf(cell, le16(data, tc))
			cell.width = les16(data, tc + 2)
			cell.borderTop = brc80(data, tc + 4)
			cell.borderLeft = brc80(data, tc + 8)
			cell.borderBottom = brc80(data, tc + 12)
			cell.borderRight = brc80(data, tc + 16)
		}

		cells.push(cell)
	}

	row.cells = cells
	row.left = cells[0]?.left ?? row.left
}

function setBorders(cell: DocCell, which: number, border: DocBorder | undefined): void {
	if ((which & 1) !== 0) cell.borderTop = border
	if ((which & 2) !== 0) cell.borderLeft = border
	if ((which & 4) !== 0) cell.borderBottom = border
	if ((which & 8) !== 0) cell.borderRight = border
}

function shadeCells(row: DocRow, data: Uint8Array, at: number, length: number, first: number): void {
	const count = Math.floor((length - 1) / 10)

	for (let k = 0; k < count; k++) {
		const cell = row.cells[first + k]

		if (cell !== undefined) cell.shading = shd(data, at + 1 + 10 * k)
	}
}

// Applies the table sprms of data[start, end) (a row mark's PAPX or a table style's TAPX) to `row`.
export function applyTable(row: DocRow, data: Uint8Array, start: number, end: number, untyped?: (sprm: number) => void): void {
	forEachPrl(data, start, end, (sprm, at, length) => {
		if (sprmSgc(sprm) !== SGC_TABLE) return

		const b = data[at] ?? 0
		const w = le16(data, at)
		const s = les16(data, at)
		const operandEnd = at + length

		switch (sprm) {
			case 0xd608:
				defTable(row, data, at, operandEnd)
				break
			case 0x5400:
			case 0x548a:
				row.jc = w
				break
			case 0x9601: {
				const shift = s - row.left

				row.left = s
				row.cells = row.cells.map(cell => ({ ...cell, left: cell.left + shift, right: cell.right + shift }))
				break
			}
			case 0x9602:
				row.gapHalf = s
				break
			case 0x9407:
				row.height = s
				break
			case 0x3403:
			case 0x3466:
				row.cantSplit = b !== 0
				break
			case 0x3404:
				row.header = b !== 0
				break
			case 0x560b:
			case 0x5664:
				row.rtl = w !== 0
				break
			case 0x3615:
				row.autofit = b !== 0
				break
			case 0x563a:
				row.istd = w
				break
			case 0xf614:
				row.width = width(data, at)
				break
			case 0xf661:
				row.indent = width(data, at)
				break
			case 0xf617:
				row.widthBefore = width(data, at)
				break
			case 0xf618:
				row.widthAfter = width(data, at)
				break
			case 0xd605:
				if (length >= 25) {
					row.borderTop = brc80(data, at + 1)
					row.borderLeft = brc80(data, at + 5)
					row.borderBottom = brc80(data, at + 9)
					row.borderRight = brc80(data, at + 13)
					row.borderInsideH = brc80(data, at + 17)
					row.borderInsideV = brc80(data, at + 21)
				}
				break
			case 0xd613:
				if (length >= 49) {
					row.borderTop = brc(data, at + 1)
					row.borderLeft = brc(data, at + 9)
					row.borderBottom = brc(data, at + 17)
					row.borderRight = brc(data, at + 25)
					row.borderInsideH = brc(data, at + 33)
					row.borderInsideV = brc(data, at + 41)
				}
				break
			case 0xd620:
				if (length >= 8) {
					const border = brc80(data, at + 4)

					forCells(row, data[at + 1] ?? 0, data[at + 2] ?? 0, cell => setBorders(cell, data[at + 3] ?? 0, border))
				}
				break
			case 0xd62f:
				if (length >= 12) {
					const border = brc(data, at + 4)

					forCells(row, data[at + 1] ?? 0, data[at + 2] ?? 0, cell => setBorders(cell, data[at + 3] ?? 0, border))
				}
				break
			case 0x7621: {
				const first = Math.min(b, row.cells.length)
				const count = Math.min(data[at + 1] ?? 0, MAX_CELLS - row.cells.length)
				const dxa = les16(data, at + 2)
				const base = row.cells[first - 1]?.right ?? row.left
				const added = Array.from({ length: Math.max(0, count) }, (_, k) => emptyCell(base + k * dxa, base + (k + 1) * dxa))
				const shift = added.length * dxa

				row.cells = [
					...row.cells.slice(0, first),
					...added,
					...row.cells.slice(first).map(cell => ({ ...cell, left: cell.left + shift, right: cell.right + shift }))
				]
				break
			}
			case 0x5622:
				if ((data[at + 1] ?? 0) > b) {
					const removed = row.cells.slice(b, data[at + 1] ?? 0)
					const shift = removed.reduce((sum, cell) => sum + (cell.right - cell.left), 0)

					row.cells = [
						...row.cells.slice(0, b),
						...row.cells.slice(data[at + 1] ?? 0).map(cell => ({ ...cell, left: cell.left - shift, right: cell.right - shift }))
					]
				}
				break
			case 0x7623: {
				const first = b
				const lim = Math.min(data[at + 1] ?? 0, row.cells.length)
				const dxa = les16(data, at + 2)
				let shift = 0

				row.cells = row.cells.map((cell, k) => {
					if (k >= first && k < lim) {
						const next = { ...cell, left: cell.left + shift, right: cell.left + shift + dxa }

						shift += dxa - (cell.right - cell.left)

						return next
					}

					return shift === 0 ? cell : { ...cell, left: cell.left + shift, right: cell.right + shift }
				})
				break
			}
			case 0x5624:
				forCells(row, b, data[at + 1] ?? 0, cell => (cell.horzMerge = 2))
				if (row.cells[b] !== undefined && (data[at + 1] ?? 0) > b) row.cells[b].horzMerge = 1
				break
			case 0x5625:
				forCells(row, b, data[at + 1] ?? 0, cell => (cell.horzMerge = 0))
				break
			case 0xd62b:
				if (length >= 3) {
					const cell = row.cells[data[at + 1] ?? 0]

					if (cell !== undefined) cell.vertMerge = (data[at + 2] ?? 0) & 3
				}
				break
			case 0xd62c:
				if (length >= 4) forCells(row, data[at + 1] ?? 0, data[at + 2] ?? 0, cell => (cell.vertAlign = data[at + 3] ?? 0))
				break
			case 0xd635:
				if (length >= 6) {
					forCells(row, data[at + 1] ?? 0, data[at + 2] ?? 0, cell => {
						cell.widthUnit = data[at + 3] ?? 0
						cell.width = les16(data, at + 4)
					})
				}
				break
			case 0xd612:
				shadeCells(row, data, at, length, 0)
				break
			case 0xd616:
				shadeCells(row, data, at, length, 22)
				break
			case 0xd60c:
				shadeCells(row, data, at, length, 44)
				break
			case 0xd634:
				// CSSA: itcFirst, itcLim, grfbrc, ftsWidth, wWidth (§2.9.48).
				if (length >= 7) {
					const sides = data[at + 3] ?? 0
					const value = le16(data, at + 5)

					for (let side = 0; side < 4; side++) if ((sides & (1 << side)) !== 0) row.padding[side] = value
				}
				break
			case 0xd633:
				if (length >= 7) row.spacing = le16(data, at + 5)
				break
			case 0x360d:
			case 0x940e:
			case 0x940f:
			case 0x9410:
			case 0x9411:
			case 0x941e:
			case 0x941f:
				row.positioned = true
				break
			case 0xd667:
				if (length >= 8) {
					row.formatChange =
						(data[at + 1] ?? 0) === 0 ? undefined : { author: les16(data, at + 2), date: dttm(le32(data, at + 4)) }
				}
				break
			default:
				untyped?.(sprm)
		}
	})
}

export interface DocColumn {
	readonly width: number
	readonly spacing: number
}

export interface DocSectionProps {
	// SBkcOperand: 0 continuous, 1 new column, 2 new page, 3 even page, 4 odd page.
	breakType: number
	titlePage: boolean
	columns: number
	columnSpacing: number
	evenlySpaced: boolean
	columnWidths: DocColumn[]
	lineBetween: boolean
	pageWidth: number
	pageHeight: number
	landscape: boolean
	marginLeft: number
	marginRight: number
	marginTop: number
	marginBottom: number
	gutter: number
	headerDistance: number
	footerDistance: number
	// MSONFC of page numbers.
	pageNumberFormat: number
	pageNumberRestart: boolean
	pageNumberStart: number
	vertAlign: number
	lineNumbering: { readonly restart: number; readonly countBy: number; readonly distance: number; readonly start: number } | undefined
	rtl: boolean
	textFlow: number
	borderTop: DocBorder | undefined
	borderLeft: DocBorder | undefined
	borderBottom: DocBorder | undefined
	borderRight: DocBorder | undefined
	pageBorderOptions: number
	gridType: number
	linePitch: number
	charSpace: number
	footnotePosition: number | undefined
	footnoteRestart: number | undefined
	footnoteStart: number | undefined
	footnoteFormat: number | undefined
	endnoteRestart: number | undefined
	endnoteStart: number | undefined
	endnoteFormat: number | undefined
	endnotesHere: boolean
	formsProtected: boolean
	formatChange: DocRevision | undefined
}

export function defaultSection(): DocSectionProps {
	return {
		breakType: 2,
		titlePage: false,
		columns: 1,
		columnSpacing: 720,
		evenlySpaced: true,
		columnWidths: [],
		lineBetween: false,
		pageWidth: 12240,
		pageHeight: 15840,
		landscape: false,
		marginLeft: 1800,
		marginRight: 1800,
		marginTop: 1440,
		marginBottom: 1440,
		gutter: 0,
		headerDistance: 720,
		footerDistance: 720,
		pageNumberFormat: 0,
		pageNumberRestart: false,
		pageNumberStart: 0,
		vertAlign: 0,
		lineNumbering: undefined,
		rtl: false,
		textFlow: 0,
		borderTop: undefined,
		borderLeft: undefined,
		borderBottom: undefined,
		borderRight: undefined,
		pageBorderOptions: 0,
		gridType: 0,
		linePitch: 0,
		charSpace: 0,
		footnotePosition: undefined,
		footnoteRestart: undefined,
		footnoteStart: undefined,
		footnoteFormat: undefined,
		endnoteRestart: undefined,
		endnoteStart: undefined,
		endnoteFormat: undefined,
		endnotesHere: false,
		formsProtected: false,
		formatChange: undefined
	}
}

function lineNumbering(target: DocSectionProps): NonNullable<DocSectionProps["lineNumbering"]> {
	return target.lineNumbering ?? { restart: 0, countBy: 0, distance: 0, start: 0 }
}

function column(target: DocSectionProps, index: number, change: Partial<DocColumn>): void {
	if (index > 44) return

	while (target.columnWidths.length <= index) target.columnWidths.push({ width: 0, spacing: 0 })

	const current = target.columnWidths[index] ?? { width: 0, spacing: 0 }

	target.columnWidths[index] = { ...current, ...change }
}

export function applySection(
	target: DocSectionProps,
	data: Uint8Array,
	start: number,
	end: number,
	untyped?: (sprm: number) => void
): void {
	forEachPrl(data, start, end, (sprm, at, length) => {
		if (sprmSgc(sprm) !== SGC_SECTION) return

		const b = data[at] ?? 0
		const w = le16(data, at)
		const s = les16(data, at)

		switch (sprm) {
			case 0x3009:
				target.breakType = b
				break
			case 0x300a:
				target.titlePage = b !== 0
				break
			case 0x500b:
				target.columns = Math.min(45, w + 1)
				break
			case 0x900c:
				target.columnSpacing = s
				break
			case 0x3005:
				target.evenlySpaced = b !== 0
				break
			case 0xf203:
				if (length >= 3) column(target, b, { width: le16(data, at + 1) })
				break
			case 0xf204:
				if (length >= 3) column(target, b, { spacing: le16(data, at + 1) })
				break
			case 0x3019:
				target.lineBetween = b !== 0
				break
			case 0xb01f:
				target.pageWidth = w
				break
			case 0xb020:
				target.pageHeight = w
				break
			case 0x301d:
				target.landscape = b === 2
				break
			case 0xb021:
				target.marginLeft = w
				break
			case 0xb022:
				target.marginRight = w
				break
			case 0x9023:
				target.marginTop = s
				break
			case 0x9024:
				target.marginBottom = s
				break
			case 0xb025:
				target.gutter = w
				break
			case 0xb017:
				target.headerDistance = w
				break
			case 0xb018:
				target.footerDistance = w
				break
			case 0x300e:
				target.pageNumberFormat = b
				break
			case 0x3011:
				target.pageNumberRestart = b !== 0
				break
			case 0x501c:
				target.pageNumberStart = w
				break
			case 0x7044:
				target.pageNumberStart = le32(data, at)
				break
			case 0x301a:
				target.vertAlign = b
				break
			case 0x3013:
				target.lineNumbering = { ...lineNumbering(target), restart: b }
				break
			case 0x5015:
				target.lineNumbering = { ...lineNumbering(target), countBy: w }
				break
			case 0x9016:
				target.lineNumbering = { ...lineNumbering(target), distance: s }
				break
			case 0x501b:
				target.lineNumbering = { ...lineNumbering(target), start: w }
				break
			case 0x3228:
				target.rtl = b !== 0
				break
			case 0x5033:
				target.textFlow = w
				break
			case 0x702b:
				target.borderTop = brc80(data, at)
				break
			case 0x702c:
				target.borderLeft = brc80(data, at)
				break
			case 0x702d:
				target.borderBottom = brc80(data, at)
				break
			case 0x702e:
				target.borderRight = brc80(data, at)
				break
			case 0xd234:
				if (length >= 9) target.borderTop = brc(data, at + 1)
				break
			case 0xd235:
				if (length >= 9) target.borderLeft = brc(data, at + 1)
				break
			case 0xd236:
				if (length >= 9) target.borderBottom = brc(data, at + 1)
				break
			case 0xd237:
				if (length >= 9) target.borderRight = brc(data, at + 1)
				break
			case 0x522f:
				target.pageBorderOptions = w
				break
			case 0x5032:
				target.gridType = w
				break
			case 0x9031:
				target.linePitch = s
				break
			case 0x7030:
				target.charSpace = les32(data, at)
				break
			case 0x303b:
				target.footnotePosition = b
				break
			case 0x303c:
				target.footnoteRestart = b
				break
			case 0x503f:
				target.footnoteStart = w
				break
			case 0x5040:
				target.footnoteFormat = w
				break
			case 0x303e:
				target.endnoteRestart = b
				break
			case 0x5041:
				target.endnoteStart = w
				break
			case 0x5042:
				target.endnoteFormat = w
				break
			case 0x3012:
				target.endnotesHere = b !== 0
				break
			case 0x3006:
				target.formsProtected = b !== 0
				break
			case 0xd243:
				if (length >= 8) {
					target.formatChange =
						(data[at + 1] ?? 0) === 0 ? undefined : { author: les16(data, at + 2), date: dttm(le32(data, at + 4)) }
				}
				break
			default:
				untyped?.(sprm)
		}
	})
}
