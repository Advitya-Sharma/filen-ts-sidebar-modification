// Paper sizes in points (1/72 in), portrait: width then height. ST_PaperSize as ECMA-376 Part 1 §18.3.1.63 lists it
// for pageSetup's paperSize (values 48 and 49 are not defined there). A size the table lacks prints on Letter, the
// attribute's default.

const INCH = 72
const MM = 72 / 25.4

function inches(width: number, height: number): readonly [number, number] {
	return [width * INCH, height * INCH]
}

function mm(width: number, height: number): readonly [number, number] {
	return [width * MM, height * MM]
}

const SIZES: ReadonlyMap<number, readonly [number, number]> = new Map([
	[1, inches(8.5, 11)],
	[2, inches(8.5, 11)],
	[3, inches(11, 17)],
	[4, inches(17, 11)],
	[5, inches(8.5, 14)],
	[6, inches(5.5, 8.5)],
	[7, inches(7.25, 10.5)],
	[8, mm(297, 420)],
	[9, mm(210, 297)],
	[10, mm(210, 297)],
	[11, mm(148, 210)],
	[12, mm(250, 353)],
	[13, mm(176, 250)],
	[14, inches(8.5, 13)],
	[15, mm(215, 275)],
	[16, inches(10, 14)],
	[17, inches(11, 17)],
	[18, inches(8.5, 11)],
	[19, inches(3.875, 8.875)],
	[20, inches(4.125, 9.5)],
	[21, inches(4.5, 10.375)],
	[22, inches(4.75, 11)],
	[23, inches(5, 11.5)],
	[24, inches(17, 22)],
	[25, inches(22, 34)],
	[26, inches(34, 44)],
	[27, mm(110, 220)],
	[28, mm(162, 229)],
	[29, mm(324, 458)],
	[30, mm(229, 324)],
	[31, mm(114, 162)],
	[32, mm(114, 229)],
	[33, mm(250, 353)],
	[34, mm(176, 250)],
	[35, mm(176, 125)],
	[36, mm(110, 230)],
	[37, inches(3.875, 7.5)],
	[38, inches(3.625, 6.5)],
	[39, inches(14.875, 11)],
	[40, inches(8.5, 12)],
	[41, inches(8.5, 13)],
	[42, mm(250, 353)],
	[43, mm(200, 148)],
	[44, inches(9, 11)],
	[45, inches(10, 11)],
	[46, inches(15, 11)],
	[47, mm(220, 220)],
	[50, inches(9.275, 12)],
	[51, inches(9.275, 15)],
	[52, inches(11.69, 18)],
	[53, mm(236, 322)],
	[54, inches(8.275, 11)],
	[55, mm(210, 297)],
	[56, inches(9.275, 12)],
	[57, mm(227, 356)],
	[58, mm(305, 487)],
	[59, inches(8.5, 12.69)],
	[60, mm(210, 330)],
	[61, mm(148, 210)],
	[62, mm(182, 257)],
	[63, mm(322, 445)],
	[64, mm(174, 235)],
	[65, mm(201, 276)],
	[66, mm(420, 594)],
	[67, mm(297, 420)],
	[68, mm(322, 445)]
])

export const LETTER = 1
export const A4 = 9

// The sizes a print dialog offers, in this order.
export const COMMON_PAPER_SIZES: readonly number[] = [1, 5, 3, 7, 9, 8, 11, 13]

export function paperSizeKnown(size: number): boolean {
	return SIZES.has(size)
}

// A size by its number, portrait, in points.
export function paperPoints(size: number): { readonly width: number; readonly height: number } {
	const [width, height] = SIZES.get(size) ?? SIZES.get(LETTER) ?? [612, 792]

	return { width, height }
}

const UNITS: Readonly<Record<string, number>> = { mm: MM, cm: 10 * MM, in: INCH, pt: 1, pc: 12, pi: 12 }

// ST_PositiveUniversalMeasure (§22.9.2.12): a positive number and a unit, in points; null when it does not read as one.
export function universalMeasure(text: string | undefined): number | null {
	if (text === undefined) return null

	const match = /^\s*([0-9]+(?:\.[0-9]+)?)(mm|cm|in|pt|pc|pi)\s*$/.exec(text)

	if (match === null) return null

	const value = Number(match[1]) * (UNITS[match[2] ?? ""] ?? Number.NaN)

	return Number.isFinite(value) && value > 0 ? value : null
}
