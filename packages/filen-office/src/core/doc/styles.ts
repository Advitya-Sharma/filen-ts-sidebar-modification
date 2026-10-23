import { type Bytes, utf16 } from "./bytes"
import { type DocChar, type DocPara, type DocRow, applyChar, applyPara, applyTable, defaultChar, defaultPara, defaultRow } from "./props"

// [MS-DOC] §2.9.271 STSH: LPStshi (§2.9.136) then rglpstd (LPStd §2.9.138 → STD §2.9.258).

export const ISTD_NIL = 0x0fff
// stk (§2.9.260 StdfBase).
export const STK_PARAGRAPH = 1
export const STK_CHARACTER = 2
export const STK_TABLE = 3
export const STK_NUMBERING = 4

export interface DocStyle {
	readonly istd: number
	readonly name: string
	// Invariant style identifier; 0x0FFE for user-defined styles.
	readonly sti: number
	readonly stk: number
	readonly istdBase: number
	readonly istdNext: number
	readonly istdLink: number
	readonly hidden: boolean
	readonly semiHidden: boolean
	readonly locked: boolean
	readonly qFormat: boolean
	readonly unhideWhenUsed: boolean
	readonly autoRedefine: boolean
	readonly priority: number | undefined
	// The style's own formatting differences (views into the Table stream).
	readonly papx: Uint8Array | undefined
	readonly chpx: Uint8Array | undefined
	readonly tapx: Uint8Array | undefined
}

export interface DocStyleSheet {
	// Indexed by istd; undefined for empty slots.
	readonly styles: readonly (DocStyle | undefined)[]
	// Stshif.ftcAsci/ftcFE/ftcOther and STSHI.ftcBi: default fonts.
	readonly fonts: { readonly ascii: number; readonly eastAsia: number; readonly other: number; readonly bi: number }
	// Stshif.stiMaxWhenSaved: how many built-in styles the saving application knew (0 when unknown).
	readonly stiMax?: number
}

const MAX_STYLES = 4096
// Inheritance deeper than this is cut (Word builds far shallower chains).
const MAX_STYLE_DEPTH = 128

function lpUpx(table: Bytes, at: number, end: number): readonly [Uint8Array | undefined, number] {
	if (at + 2 > end) return [undefined, end]

	const cb = table.u16(at)

	if (at + 2 + cb > end) return [undefined, end]

	return [table.data.subarray(at + 2, at + 2 + cb), at + 2 + cb + (cb & 1)]
}

function readStd(table: Bytes, istd: number, at: number, end: number, cbStdBase: number): DocStyle | undefined {
	if (at + 10 > end) return undefined

	const w0 = table.u16(at)
	const w1 = table.u16(at + 2)
	const w2 = table.u16(at + 4)
	const grfstd = table.u16(at + 8)
	const post2000 = cbStdBase >= 18 && at + 18 <= end
	const w5 = post2000 ? table.u16(at + 10) : 0
	const w8 = post2000 ? table.u16(at + 16) : 0
	const stk = w1 & 0xf
	const cupx = w2 & 0xf
	let cursor = at + cbStdBase
	let name = ""

	// xstzName: Xst (cch, chars) + chTerm.
	if (cursor + 2 <= end) {
		const cch = table.u16(cursor)

		if (cursor + 2 + 2 * cch <= end) {
			name = utf16(table.data, cursor + 2, cch)
			cursor += 2 + 2 * cch + 2
		} else {
			cursor = end
		}
	}

	if ((cursor - at) & 1) cursor++

	const upx: (Uint8Array | undefined)[] = []

	for (let k = 0; k < cupx && cursor < end; k++) {
		const [bytes, next] = lpUpx(table, cursor, end)

		upx.push(bytes)
		cursor = next
	}

	let papx: Uint8Array | undefined
	let chpx: Uint8Array | undefined
	let tapx: Uint8Array | undefined

	// StkParaGRLPUPX: papx (istd + grpprl), chpx; StkCharGRLPUPX: chpx; StkTableGRLPUPX: tapx, papx, chpx;
	// StkListGRLPUPX: papx.
	const papxGrpprl = (bytes: Uint8Array | undefined): Uint8Array | undefined =>
		bytes !== undefined && bytes.length >= 2 ? bytes.subarray(2) : undefined

	if (stk === STK_PARAGRAPH) {
		papx = papxGrpprl(upx[0])
		chpx = upx[1]
	} else if (stk === STK_CHARACTER) {
		chpx = upx[0]
	} else if (stk === STK_TABLE) {
		tapx = upx[0]
		papx = papxGrpprl(upx[1])
		chpx = upx[2]
	} else if (stk === STK_NUMBERING) {
		papx = papxGrpprl(upx[0])
	}

	return {
		istd,
		name,
		sti: w0 & 0xfff,
		stk,
		istdBase: w1 >> 4,
		istdNext: w2 >> 4,
		istdLink: w5 & 0xfff,
		autoRedefine: (grfstd & 0x0001) !== 0,
		hidden: (grfstd & 0x0002) !== 0,
		semiHidden: (grfstd & 0x0100) !== 0,
		locked: (grfstd & 0x0200) !== 0,
		unhideWhenUsed: (grfstd & 0x0800) !== 0,
		qFormat: (grfstd & 0x1000) !== 0,
		priority: post2000 ? (w8 >> 4) & 0xfff : undefined,
		papx,
		chpx,
		tapx
	}
}

export function readStyles(table: Bytes, fc: number, lcb: number): DocStyleSheet {
	const empty: DocStyleSheet = { styles: [], fonts: { ascii: 0, eastAsia: 0, other: 0, bi: 0 } }

	if (lcb < 2 || !table.has(fc, lcb)) return empty

	const end = fc + lcb
	const cbStshi = table.u16(fc)
	const stshi = fc + 2

	if (cbStshi < 18 || stshi + cbStshi > end) return empty

	const cstd = Math.min(MAX_STYLES, table.u16(stshi))
	const cbStdBase = table.u16(stshi + 2)
	const fonts = {
		ascii: table.u16(stshi + 12),
		eastAsia: table.u16(stshi + 14),
		other: table.u16(stshi + 16),
		bi: cbStshi >= 20 ? table.u16(stshi + 18) : 0
	}
	const styles: (DocStyle | undefined)[] = []
	let at = stshi + cbStshi

	for (let istd = 0; istd < cstd && at + 2 <= end; istd++) {
		const cbStd = table.u16(at)
		const stdEnd = Math.min(end, at + 2 + cbStd)

		styles.push(cbStd === 0 ? undefined : readStd(table, istd, at + 2, stdEnd, cbStdBase))
		at = at + 2 + cbStd + (cbStd & 1)
	}

	return { styles, fonts, stiMax: table.u16(stshi + 6) }
}

function copyRow(row: DocRow): DocRow {
	return { ...row, cells: [], padding: [...row.padding] }
}

export interface ResolvedStyle {
	readonly para: DocPara
	readonly char: DocChar
	readonly row: DocRow | undefined
}

// Resolves styles through istdBase (§2.4.6.5), memoised; chains are walked iteratively and loops cut.
export class StyleResolver {
	readonly sheet: DocStyleSheet
	private readonly defaults: DocChar
	private readonly cache = new Map<number, ResolvedStyle>()
	private readonly charCache = new Map<string, DocChar>()

	constructor(sheet: DocStyleSheet) {
		this.sheet = sheet
		this.defaults = defaultChar()
		this.defaults.fontAscii = sheet.fonts.ascii
		this.defaults.fontEastAsia = sheet.fonts.eastAsia
		this.defaults.fontOther = sheet.fonts.other
		this.defaults.fontBi = sheet.fonts.bi
	}

	get defaultChar(): DocChar {
		return this.defaults
	}

	style(istd: number): DocStyle | undefined {
		return this.sheet.styles[istd]
	}

	// The styles from `istd` up its istdBase chain, base first; loops are cut and the walk stops below `stop`.
	private chain(istd: number, stop: (istd: number) => boolean = () => false): DocStyle[] {
		const out: DocStyle[] = []
		const seen = new Set<number>()

		for (let current = istd; current !== ISTD_NIL && !seen.has(current) && out.length < MAX_STYLE_DEPTH;) {
			const style = this.sheet.styles[current]

			if (style === undefined) break

			seen.add(current)
			out.push(style)

			if (stop(style.istdBase)) break

			current = style.istdBase
		}

		return out.reverse()
	}

	// Memoised along the chain, so resolving every style is linear in the sheet.
	resolve(istd: number): ResolvedStyle {
		const cached = this.cache.get(istd)

		if (cached !== undefined) return cached

		const chain = this.chain(istd, base => this.cache.has(base))
		const base = chain.length > 0 ? this.cache.get(chain[0]?.istdBase ?? ISTD_NIL) : undefined
		let para = base === undefined ? defaultPara() : { ...base.para, tabs: [...base.para.tabs] }
		let char = base === undefined ? { ...this.defaults } : { ...base.char }
		let row: DocRow | undefined = base?.row === undefined ? undefined : copyRow(base.row)
		let resolved: ResolvedStyle = { para, char, row }

		for (const style of chain) {
			para = { ...para, tabs: [...para.tabs] }
			char = { ...char }
			row = row === undefined ? undefined : copyRow(row)

			if (style.papx !== undefined) applyPara(para, style.papx, 0, style.papx.length)

			if (style.chpx !== undefined) applyChar(char, style.chpx, 0, style.chpx.length, { style: { ...char } })

			if (style.tapx !== undefined) {
				row ??= defaultRow()
				applyTable(row, style.tapx, 0, style.tapx.length)
			}

			para.istd = style.istd
			resolved = { para, char, row }
			this.cache.set(style.istd, resolved)
		}

		if (chain.length === 0) {
			para.istd = istd
			this.cache.set(istd, resolved)
		}

		return resolved
	}

	// The character properties of text in a paragraph of style `paraIstd` with character style `charIstd` (§2.4.6.6
	// part 2 step 4, then sprmCIstd).
	charBase(paraIstd: number, charIstd: number): DocChar {
		const key = `${paraIstd}:${charIstd}`
		const cached = this.charCache.get(key)

		if (cached !== undefined) return cached

		const char = { ...this.resolve(paraIstd).char }

		if (charIstd !== 10 && this.sheet.styles[charIstd]?.stk === STK_CHARACTER) {
			for (const style of this.chain(charIstd)) {
				if (style.stk === STK_CHARACTER && style.chpx !== undefined)
					applyChar(char, style.chpx, 0, style.chpx.length, { style: { ...char } })
			}
		}

		char.istd = charIstd

		this.charCache.set(key, char)

		return char
	}
}
