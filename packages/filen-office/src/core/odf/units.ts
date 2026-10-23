// Value types of ODF 1.3 Part 3 §18.3: lengths, percents, relative lengths, angles, colors, booleans.
// Invalid values read as undefined, never as 0: a caller tells "absent" from "zero".

const POINTS_PER_UNIT: Readonly<Record<string, number>> = {
	pt: 1,
	pc: 12,
	in: 72,
	// OpenOffice.org 2.x wrote "inch"; LibreOffice still reads it.
	inch: 72,
	cm: 72 / 2.54,
	mm: 72 / 25.4,
	// CSS pixels (1/96 in), as LibreOffice reads them.
	px: 0.75
}

// §18.3.18 length: -?([0-9]+(\.[0-9]*)?|\.[0-9]+)(cm|mm|in|pt|pc|px).
const LENGTH = /^\s*(-?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?)\s*(cm|mm|inch|in|pt|pc|px)\s*$/
// §18.3.23 percent.
const PERCENT = /^\s*(-?(?:\d+(?:\.\d*)?|\.\d+))%\s*$/
// §18.3.27 relativeLength: [0-9]+\*.
const RELATIVE = /^\s*(\d+(?:\.\d*)?)\*\s*$/
// §18.3.1 angle: a double with an optional unit (degrees when none).
const ANGLE = /^\s*(-?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?)\s*(deg|grad|rad)?\s*$/
// §18.3.8 color.
const COLOR = /^#([0-9a-fA-F]{6})$/

// A length in points (1 pt = 1/72 in).
export function parseLength(value: string | undefined): number | undefined {
	if (value === undefined) return undefined

	const match = LENGTH.exec(value)

	if (match === null) return undefined

	const number = Number(match[1])
	const scale = POINTS_PER_UNIT[match[2] ?? ""]

	return Number.isFinite(number) && scale !== undefined ? number * scale : undefined
}

// A percent as written (50% → 50).
export function parsePercent(value: string | undefined): number | undefined {
	if (value === undefined) return undefined

	const match = PERCENT.exec(value)

	return match === null ? undefined : Number(match[1])
}

export type LengthOrPercent = { readonly kind: "length"; readonly points: number } | { readonly kind: "percent"; readonly percent: number }

// Attributes typed "length or percent" (fo:font-size, fo:margin-left, fo:line-height...).
export function parseLengthOrPercent(value: string | undefined): LengthOrPercent | undefined {
	const points = parseLength(value)

	if (points !== undefined) return { kind: "length", points }

	const percent = parsePercent(value)

	return percent === undefined ? undefined : { kind: "percent", percent }
}

// A relative length's weight ("3*" → 3), as style:rel-width and style:rel-column-width use.
export function parseRelativeLength(value: string | undefined): number | undefined {
	if (value === undefined) return undefined

	const match = RELATIVE.exec(value)

	return match === null ? undefined : Number(match[1])
}

// An angle in degrees.
export function parseAngle(value: string | undefined): number | undefined {
	if (value === undefined) return undefined

	const match = ANGLE.exec(value)

	if (match === null) return undefined

	const number = Number(match[1])

	if (!Number.isFinite(number)) return undefined

	switch (match[2]) {
		case "grad":
			return (number * 360) / 400
		case "rad":
			return (number * 180) / Math.PI
		default:
			return number
	}
}

// A color as 0xRRGGBB.
export function parseColor(value: string | undefined): number | undefined {
	if (value === undefined) return undefined

	const match = COLOR.exec(value.trim())

	return match === null ? undefined : parseInt(match[1] ?? "0", 16)
}

// §18.3.2 boolean: exactly "true" or "false".
export function parseBoolean(value: string | undefined): boolean | undefined {
	return value === "true" ? true : value === "false" ? false : undefined
}

export function pointsToTwips(points: number): number {
	return Math.round(points * 20)
}

// CSS pixels at 96 dpi, the unit the layouts use.
export function pointsToPixels(points: number): number {
	return (points * 96) / 72
}
