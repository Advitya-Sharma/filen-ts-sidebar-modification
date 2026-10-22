// Header and footer text (ECMA-376 Part 1 §18.3.1.46 oddHeader and its siblings; Excel's codes as "Formatting and
// VBA codes for headers and footers" lists them on learn.microsoft.com): three sections (&L, &C, &R) of runs, each run
// plain text or a field filled in per page, with the run's formatting toggled by codes.

export type HeaderFooterField = "page" | "pages" | "date" | "time" | "file" | "sheet" | "path" | "picture"

export interface HeaderFooterStyle {
	readonly bold: boolean
	readonly italic: boolean
	readonly underline: "single" | "double" | null
	readonly strike: boolean
	readonly script: "superscript" | "subscript" | null
	// A family the text names (&"name,style"), else the sheet's.
	readonly font: string | null
	// Points (&nn), else the default (11).
	readonly size: number | null
	// 0xRRGGBB (&Krrggbb), else automatic.
	readonly color: number | null
}

export type HeaderFooterRun =
	| { readonly kind: "text"; readonly text: string; readonly style: HeaderFooterStyle }
	// &P+n and &P-n: the page number plus `offset`.
	| { readonly kind: "field"; readonly field: HeaderFooterField; readonly offset: number; readonly style: HeaderFooterStyle }

export interface HeaderFooterSections {
	readonly left: readonly HeaderFooterRun[]
	readonly center: readonly HeaderFooterRun[]
	readonly right: readonly HeaderFooterRun[]
}

const PLAIN: HeaderFooterStyle = {
	bold: false,
	italic: false,
	underline: null,
	strike: false,
	script: null,
	font: null,
	size: null,
	color: null
}

const FIELDS: Readonly<Record<string, HeaderFooterField>> = {
	P: "page",
	N: "pages",
	D: "date",
	T: "time",
	F: "file",
	A: "sheet",
	Z: "path",
	G: "picture"
}

export const EMPTY_SECTIONS: HeaderFooterSections = { left: [], center: [], right: [] }

// Text before any section code is the center section's, as Excel reads it.
export function parseHeaderFooter(text: string): HeaderFooterSections {
	const sections: Record<"left" | "center" | "right", HeaderFooterRun[]> = { left: [], center: [], right: [] }
	let current: "left" | "center" | "right" = "center"
	let style: HeaderFooterStyle = PLAIN
	let pending = ""

	const flush = (): void => {
		if (pending === "") return

		const runs = sections[current]
		const last = runs[runs.length - 1]

		if (last?.kind === "text" && last.style === style) runs[runs.length - 1] = { kind: "text", text: last.text + pending, style }
		else runs.push({ kind: "text", text: pending, style })

		pending = ""
	}
	const restyle = (next: HeaderFooterStyle): void => {
		flush()
		style = next
	}

	let i = 0

	while (i < text.length) {
		const char = text[i] ?? ""

		if (char !== "&" || i + 1 >= text.length) {
			pending += char
			i++

			continue
		}

		const code = text[i + 1] ?? ""
		const upper = code.toUpperCase()

		i += 2

		if (code === "&") {
			pending += "&"
		} else if (upper === "L" || upper === "C" || upper === "R") {
			flush()
			current = upper === "L" ? "left" : upper === "C" ? "center" : "right"
			// Each section starts plain, as Excel shows it.
			style = PLAIN
		} else if (FIELDS[upper] !== undefined) {
			const field = FIELDS[upper]
			let offset = 0

			if (field === "page") {
				const match = /^[+-][0-9]{1,5}/.exec(text.slice(i))

				if (match !== null) {
					offset = Number(match[0])
					i += match[0].length
				}
			}

			flush()
			sections[current].push({ kind: "field", field, offset, style })
		} else if (upper === "B") {
			restyle({ ...style, bold: !style.bold })
		} else if (upper === "I") {
			restyle({ ...style, italic: !style.italic })
		} else if (upper === "U") {
			restyle({ ...style, underline: style.underline === "single" ? null : "single" })
		} else if (upper === "E") {
			restyle({ ...style, underline: style.underline === "double" ? null : "double" })
		} else if (upper === "S") {
			restyle({ ...style, strike: !style.strike })
		} else if (upper === "X") {
			restyle({ ...style, script: style.script === "superscript" ? null : "superscript" })
		} else if (upper === "Y") {
			restyle({ ...style, script: style.script === "subscript" ? null : "subscript" })
		} else if (code === '"') {
			const end = text.indexOf('"', i)
			const spec = end < 0 ? text.slice(i) : text.slice(i, end)

			i = end < 0 ? text.length : end + 1

			const [family = "", face = ""] = spec.split(",")
			const lowerFace = face.toLowerCase()
			const font = family === "-" || family === "+" || family.trim() === "" ? style.font : family.trim()

			restyle({
				...style,
				font,
				bold: lowerFace.includes("bold") || (face === "" && style.bold),
				italic: lowerFace.includes("italic") || (face === "" && style.italic)
			})
		} else if (/[0-9]/.test(code)) {
			const match = /^[0-9]{1,3}/.exec(text.slice(i - 1))
			const digits = match?.[0] ?? code

			i += digits.length - 1

			const size = Number(digits)

			restyle({ ...style, size: size > 0 && size <= 409 ? size : style.size })
		} else if (upper === "K") {
			const hex = /^[0-9A-Fa-f]{6}/.exec(text.slice(i))
			const theme = /^[0-9]{2}[+-][0-9]{3}/.exec(text.slice(i))

			if (hex !== null) {
				i += 6
				restyle({ ...style, color: Number.parseInt(hex[0], 16) })
			} else if (theme !== null) {
				// A theme colour: drawn automatic (the print root carries no theme).
				i += theme[0].length
				restyle({ ...style, color: null })
			}
		}
		// Any other code (&O outline, &H shadow, one this reader does not know) changes nothing drawn.
	}

	flush()

	return sections
}

export interface HeaderFooterValues {
	readonly page: number
	readonly pages: number
	readonly date: string
	readonly time: string
	readonly file: string
	readonly sheet: string
	readonly path: string
}

// A run's text on one page (a picture prints nothing: the file's header images are not drawn).
export function runText(run: HeaderFooterRun, values: HeaderFooterValues): string {
	if (run.kind === "text") return run.text

	switch (run.field) {
		case "page":
			return String(values.page + run.offset)
		case "pages":
			return String(values.pages)
		case "date":
			return values.date
		case "time":
			return values.time
		case "file":
			return values.file
		case "sheet":
			return values.sheet
		case "path":
			return values.path
		case "picture":
			return ""
	}
}

export function sectionsEmpty(sections: HeaderFooterSections): boolean {
	return sections.left.length === 0 && sections.center.length === 0 && sections.right.length === 0
}
