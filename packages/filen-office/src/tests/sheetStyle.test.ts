import { describe, expect, it } from "vitest"
import { applyTint, DEFAULT_PALETTE, resolveColor, SheetStyles, workbookPalette, type SheetStyleSource } from "../core/sheet/style"
import { NUMFMT_EN_US, formatValue, IMPLIED_BUILTINS } from "../core/sheet/numfmt"
import { XfApply, type XlsxFill, type XlsxTheme, type XlsxXf } from "../core/xlsx/model/types"

// Office's default theme colours in clrScheme order (dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink).
const THEME: XlsxTheme = {
	colors: [0x000000, 0xffffff, 0x44546a, 0xe7e6e6, 0x4472c4, 0xed7d31, 0xa5a5a5, 0xffc000, 0x5b9bd5, 0x70ad47, 0x0563c1, 0x954f72],
	majorFont: "Calibri Light",
	minorFont: "Calibri"
}

function xf(fields: Partial<XlsxXf>): XlsxXf {
	return { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0, flags: 0, ...fields }
}

function source(fields: Partial<SheetStyleSource>): SheetStyleSource {
	return {
		numFmts: new Map(),
		fonts: [{ name: "Calibri", size: 11, scheme: "minor", color: { theme: 1 } }],
		fills: [
			{ kind: "pattern", pattern: "none" },
			{ kind: "pattern", pattern: "gray125" }
		],
		borders: [{}],
		cellStyleXfs: [xf({ xfId: -1 })],
		cellXfs: [xf({})],
		indexedColors: null,
		...fields
	}
}

const hex = (rgb: number): string => rgb.toString(16).padStart(6, "0")

describe("colours", () => {
	it("read theme 0 as lt1 and 1 as dk1, the rest in scheme order", () => {
		expect(resolveColor({ theme: 0 }, THEME, DEFAULT_PALETTE)).toBe(0xffffff)
		expect(resolveColor({ theme: 1 }, THEME, DEFAULT_PALETTE)).toBe(0x000000)
		expect(resolveColor({ theme: 2 }, THEME, DEFAULT_PALETTE)).toBe(0xe7e6e6)
		expect(resolveColor({ theme: 3 }, THEME, DEFAULT_PALETTE)).toBe(0x44546a)
		expect(resolveColor({ theme: 4 }, THEME, DEFAULT_PALETTE)).toBe(0x4472c4)
		expect(resolveColor({ theme: 11 }, THEME, DEFAULT_PALETTE)).toBe(0x954f72)
		expect(resolveColor({ theme: 12 }, THEME, DEFAULT_PALETTE)).toBe(-1)
	})

	it("darken and lighten the luminance by the tint (ECMA-376 Part 1 §18.8.3)", () => {
		expect(hex(applyTint(0xffffff, -0.0499893185216834))).toBe("f2f2f2")
		expect(hex(applyTint(0xffffff, -0.249977111117893))).toBe("bfbfbf")
		expect(hex(applyTint(0xffffff, -0.499984740745262))).toBe("808080")
		expect(hex(applyTint(0x000000, 0.499984740745262))).toBe("7f7f7f")
		expect(hex(applyTint(0x000000, 0.14996795556505021))).toBe("262626")
		expect(applyTint(0x4472c4, 0)).toBe(0x4472c4)
		expect(applyTint(0x4472c4, 1)).toBe(0xffffff)
		expect(applyTint(0x4472c4, -1)).toBe(0x000000)
		// Hue and saturation stay: lightening accent 1 keeps it blue.
		const light = applyTint(0x4472c4, 0.6)

		expect(light & 0xff).toBeGreaterThan((light >> 16) & 0xff)
		expect(resolveColor({ theme: 0, tint: -0.249977111117893 }, THEME, DEFAULT_PALETTE)).toBe(0xbfbfbf)
	})

	it("take theme over rgb over indexed over auto (MS-OI29500 §2.1.694)", () => {
		expect(resolveColor({ theme: 4, argb: 0xffff0000, indexed: 3, auto: true }, THEME, DEFAULT_PALETTE)).toBe(0x4472c4)
		expect(resolveColor({ argb: 0xffff0000, indexed: 3 }, THEME, DEFAULT_PALETTE)).toBe(0xff0000)
		expect(resolveColor({ indexed: 3, auto: true }, THEME, DEFAULT_PALETTE)).toBe(0x00ff00)
		expect(resolveColor({ auto: true }, THEME, DEFAULT_PALETTE, 0x123456)).toBe(0x123456)
		expect(resolveColor(undefined, THEME, DEFAULT_PALETTE)).toBe(-1)
		expect(resolveColor({ argb: 0x80aabbcc }, THEME, DEFAULT_PALETTE)).toBe(0xaabbcc)
	})

	it("index the default palette (§18.8.27) or the workbook's own", () => {
		expect(DEFAULT_PALETTE).toHaveLength(66)
		expect(resolveColor({ indexed: 10 }, null, DEFAULT_PALETTE)).toBe(0xff0000)
		expect(resolveColor({ indexed: 22 }, null, DEFAULT_PALETTE)).toBe(0xc0c0c0)
		expect(resolveColor({ indexed: 63 }, null, DEFAULT_PALETTE)).toBe(0x333333)
		expect(resolveColor({ indexed: 64 }, null, DEFAULT_PALETTE)).toBe(0x000000)
		expect(resolveColor({ indexed: 65 }, null, DEFAULT_PALETTE)).toBe(0xffffff)
		expect(resolveColor({ indexed: 200 }, null, DEFAULT_PALETTE)).toBe(-1)

		const palette = workbookPalette([0xff112233, 0xff445566])

		expect(resolveColor({ indexed: 1 }, null, palette)).toBe(0x445566)
		expect(resolveColor({ indexed: 2 }, null, palette)).toBe(0xff0000)
		expect(workbookPalette(null)).toBe(DEFAULT_PALETTE)
	})

	it("fall back to the default theme without a theme part", () => {
		expect(resolveColor({ theme: 4 }, null, DEFAULT_PALETTE)).toBe(0x4f81bd)
	})
})

describe("render styles", () => {
	it("resolve the font, the theme's scheme font and colour", () => {
		const styles = new SheetStyles(
			source({
				fonts: [
					{ name: "Calibri", size: 11, scheme: "minor", color: { theme: 1 } },
					{
						name: "Arial",
						size: 14,
						bold: true,
						italic: true,
						strike: true,
						underline: "double",
						vertAlign: "superscript",
						color: { argb: 0xffc00000 }
					},
					{ name: "Cambria", size: 18, scheme: "major", color: { auto: true }, underline: "none" }
				],
				cellXfs: [xf({}), xf({ fontId: 1, flags: XfApply.FONT }), xf({ fontId: 2, flags: XfApply.FONT })]
			}),
			{ ...THEME, minorFont: "Aptos" }
		)

		expect(styles.renderStyle(0)).toEqual({ font: "Aptos", size: 11, color: 0x000000 })
		expect(styles.renderStyle(1)).toEqual({
			font: "Arial",
			size: 14,
			bold: true,
			italic: true,
			strike: true,
			underline: "double",
			script: "superscript",
			color: 0xc00000
		})
		expect(styles.renderStyle(2)).toEqual({ font: "Calibri Light", size: 18 })
	})

	it("use the cell format's own font and fill whatever its apply flags (MS-OI29500 §2.1.721)", () => {
		const styles = new SheetStyles(
			source({
				fonts: [
					{ name: "Calibri", size: 11 },
					{ name: "Arial", size: 9 }
				],
				fills: [
					{ kind: "pattern", pattern: "none" },
					{ kind: "pattern", pattern: "gray125" },
					{ kind: "pattern", pattern: "solid", fg: { argb: 0xffffff00 } }
				],
				cellStyleXfs: [xf({ xfId: -1 }), xf({ xfId: -1, fontId: 1, fillId: 0 })],
				cellXfs: [xf({}), xf({ xfId: 1, fontId: 0, fillId: 2 })]
			}),
			THEME
		)

		expect(styles.renderStyle(1)).toEqual({ font: "Calibri", size: 11, fill: 0xffff00 })
	})

	it("inherit alignment and protection from the cell style unless the cell has its own", () => {
		const styles = new SheetStyles(
			source({
				cellStyleXfs: [
					xf({ xfId: -1 }),
					xf({ xfId: -1, alignment: { horizontal: "center", wrapText: true }, protection: { locked: false } })
				],
				cellXfs: [
					xf({}),
					xf({ xfId: 1 }),
					xf({ xfId: 1, flags: XfApply.ALIGNMENT | XfApply.PROTECTION }),
					xf({ xfId: 1, alignment: { horizontal: "right" }, protection: { hidden: true } })
				]
			}),
			THEME
		)

		expect(styles.renderStyle(1)).toMatchObject({ align: "center", wrap: true })
		expect(styles.locked(1)).toBe(false)
		expect(styles.renderStyle(2).align).toBeUndefined()
		expect(styles.locked(2)).toBe(true)
		expect(styles.renderStyle(3)).toMatchObject({ align: "right" })
		expect(styles.renderStyle(3).wrap).toBeUndefined()
		expect(styles.locked(3)).toBe(true)
		expect(styles.formulaHidden(3)).toBe(true)
		expect(styles.locked(0)).toBe(true)
	})

	it("resolve alignment (§18.8.1)", () => {
		const styles = new SheetStyles(
			source({
				cellXfs: [
					xf({ alignment: { horizontal: "general", vertical: "bottom" } }),
					xf({
						alignment: {
							horizontal: "distributed",
							vertical: "top",
							indent: 2,
							shrinkToFit: true,
							textRotation: 45,
							readingOrder: 2
						}
					}),
					xf({ alignment: { textRotation: 135 } }),
					xf({ alignment: { textRotation: 255 } }),
					xf({ alignment: { textRotation: 90 } })
				]
			}),
			THEME
		)
		const without = ({ font: _font, size: _size, color: _color, ...rest }: Record<string, unknown>): Record<string, unknown> => rest

		expect(without({ ...styles.renderStyle(0) })).toEqual({})
		expect(without({ ...styles.renderStyle(1) })).toEqual({
			align: "distributed",
			valign: "top",
			indent: 2,
			shrink: true,
			rotation: 45,
			rightToLeft: true
		})
		expect(styles.renderStyle(2).rotation).toBe(-45)
		expect(styles.renderStyle(3).rotation).toBe(255)
		expect(styles.renderStyle(4).rotation).toBe(90)
	})

	it("resolve borders, diagonals and automatic colours", () => {
		const styles = new SheetStyles(
			source({
				borders: [
					{},
					{
						left: { style: "thin", color: { auto: true } },
						right: { style: "medium", color: { indexed: 10 } },
						top: { style: "dashed", color: { theme: 4, tint: 0.5 } },
						bottom: { style: "double" },
						diagonal: { style: "hair", color: { argb: 0xff00ff00 } },
						diagonalUp: true
					},
					{ left: {}, diagonal: { style: "thin" }, diagonalUp: true, diagonalDown: true }
				],
				cellXfs: [xf({}), xf({ borderId: 1 }), xf({ borderId: 2 })]
			}),
			THEME
		)
		const one = styles.renderStyle(1)

		expect(one.left).toEqual({ style: "thin", color: 0 })
		expect(one.right).toEqual({ style: "medium", color: 0xff0000 })
		expect(one.top).toEqual({ style: "dashed", color: applyTint(0x4472c4, 0.5) })
		expect(one.bottom).toEqual({ style: "double", color: 0 })
		expect(one.diagonalUp).toEqual({ style: "hair", color: 0x00ff00 })
		expect(one.diagonalDown).toBeUndefined()

		const two = styles.renderStyle(2)

		expect(two.left).toBeUndefined()
		expect(two.diagonalUp).toEqual({ style: "thin", color: 0 })
		expect(two.diagonalDown).toEqual({ style: "thin", color: 0 })
	})

	it("resolve solid, patterned and gradient fills (§18.8.20)", () => {
		const fills: XlsxFill[] = [
			{ kind: "pattern", pattern: "none" },
			{ kind: "pattern", pattern: "gray125" },
			{ kind: "pattern", pattern: "solid", fg: { theme: 0, tint: -0.0499893185216834 }, bg: { indexed: 64 } },
			{ kind: "pattern", pattern: "darkTrellis", fg: { argb: 0xffff0000 }, bg: { argb: 0xff0000ff } },
			{ kind: "pattern", pattern: "lightGrid" },
			{
				kind: "gradient",
				type: "linear",
				degree: 90,
				left: 0,
				right: 0,
				top: 0,
				bottom: 0,
				stops: [
					{ position: 0, color: { argb: 0xff000000 } },
					{ position: 1, color: { argb: 0xffffffff } }
				]
			},
			{
				kind: "gradient",
				type: "path",
				degree: 0,
				left: 0.5,
				right: 0.5,
				top: 0.5,
				bottom: 0.5,
				stops: [{ position: 0.5, color: { argb: 0xff336699 } }]
			}
		]
		const styles = new SheetStyles(source({ fills, cellXfs: fills.map((_, fillId) => xf({ fillId })) }), THEME)

		expect(styles.renderStyle(0).fill).toBeUndefined()
		expect(styles.renderStyle(1).pattern).toEqual({ type: "gray125", color: 0, background: 0xffffff })
		expect(styles.renderStyle(2).fill).toBe(0xf2f2f2)
		expect(styles.renderStyle(3).pattern).toEqual({ type: "darkTrellis", color: 0xff0000, background: 0x0000ff })
		expect(styles.renderStyle(4).pattern).toEqual({ type: "lightGrid", color: 0, background: 0xffffff })
		expect(styles.renderStyle(5).fill).toBe(0x808080)
		expect(styles.renderStyle(6).fill).toBe(0x336699)
	})

	it("are cached per style id and read a missing id as the first", () => {
		const styles = new SheetStyles(source({ cellXfs: [xf({}), xf({ fillId: 1 })] }), THEME)

		expect(styles.renderStyle(1)).toBe(styles.renderStyle(1))
		expect(styles.renderStyle(7)).toBe(styles.renderStyle(0))
		expect(styles.renderStyle(-1)).toBe(styles.renderStyle(0))
		expect(new SheetStyles(source({ cellXfs: [] }), null).renderStyle(0)).toEqual({})
	})
})

describe("number formats of a cell format", () => {
	it("take the workbook's record before the built-in one (§18.8.30)", () => {
		const styles = new SheetStyles(
			source({
				numFmts: new Map([
					[14, "0.00%"],
					[164, '"£"#,##0.00']
				]),
				cellXfs: [
					xf({}),
					xf({ numFmtId: 14 }),
					xf({ numFmtId: 164 }),
					xf({ numFmtId: 22 }),
					xf({ numFmtId: 30 }),
					xf({ numFmtId: 49 })
				]
			}),
			THEME
		)

		expect(styles.numFmtCode(0)).toBe("General")
		expect(styles.numFmtCode(1)).toBe("0.00%")
		expect(styles.numFmtCode(2)).toBe('"£"#,##0.00')
		expect(styles.numFmtCode(3)).toBe("m/d/yyyy h:mm")
		expect(styles.numFmtCode(4)).toBe("General")
		expect(styles.numFmt(5).isText).toBe(true)
		expect(styles.numFmt(2)).toBe(styles.numFmt(2))
	})

	it("take language-dependent built-ins from the locale", () => {
		const cellXfs = [xf({}), xf({ numFmtId: 30 }), xf({ numFmtId: 7 })]
		const ja = new SheetStyles(source({ cellXfs }), THEME, { locale: { ...NUMFMT_EN_US, builtins: IMPLIED_BUILTINS["ja-JP"] } })

		expect(ja.numFmtCode(1)).toBe("m/d/yy")
		expect(formatValue(1234.5, new SheetStyles(source({ cellXfs }), THEME).numFmt(2)).text).toBe("$1,234.50 ")
	})

	it("mark a quote prefix", () => {
		const styles = new SheetStyles(source({ cellXfs: [xf({}), xf({ flags: XfApply.QUOTE_PREFIX })] }), THEME)

		expect(styles.quotePrefix(0)).toBe(false)
		expect(styles.quotePrefix(1)).toBe(true)
	})
})
