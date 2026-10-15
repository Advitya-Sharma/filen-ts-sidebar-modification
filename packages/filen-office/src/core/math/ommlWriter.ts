import { escapeAttribute, escapeText } from "../xml/index"

// Office Math (ECMA-376 Part 1 §22.1) as plain data, and its XML: what the LaTeX reader builds and an equation
// the editor makes is written as. Elements carry the m: prefix the story root declares; nothing is written
// deeper than MAX_OMML_NESTING objects.

export type OmmlContent = readonly OmmlNode[]

// m:sty: plain, bold, italic, bold italic.
export type OmmlStyle = "p" | "b" | "i" | "bi"
// m:scr.
export type OmmlScript = "roman" | "script" | "fraktur" | "double-struck" | "sans-serif" | "monospace"

export type OmmlNode =
	// m:r; `normal` is m:nor (text, not math: LaTeX's \text).
	| {
			readonly kind: "run"
			readonly text: string
			readonly style?: OmmlStyle
			readonly script?: OmmlScript
			readonly normal?: boolean
	  }
	// m:f.
	| {
			readonly kind: "fraction"
			readonly type?: "bar" | "skw" | "lin" | "noBar"
			readonly numerator: OmmlContent
			readonly denominator: OmmlContent
	  }
	// m:rad; a null degree is m:degHide.
	| { readonly kind: "radical"; readonly degree: OmmlContent | null; readonly base: OmmlContent }
	// m:sSub, m:sSup, m:sSubSup, or m:sPre when `pre`.
	| {
			readonly kind: "script"
			readonly base: OmmlContent
			readonly sub: OmmlContent | null
			readonly sup: OmmlContent | null
			readonly pre?: boolean
	  }
	// m:nary; null limits are m:subHide/m:supHide, a null location the document's default.
	| {
			readonly kind: "nary"
			readonly operator: string
			readonly location: "undOvr" | "subSup" | null
			readonly sub: OmmlContent | null
			readonly sup: OmmlContent | null
			readonly base: OmmlContent
	  }
	// m:d, one m:e per item; "" for an absent fence.
	| {
			readonly kind: "delimiter"
			readonly open: string
			readonly close: string
			readonly separator?: string
			readonly items: readonly OmmlContent[]
	  }
	// m:m, rows of cells.
	| { readonly kind: "matrix"; readonly rows: readonly (readonly OmmlContent[])[] }
	// m:eqArr (LaTeX's cases inside a delimiter).
	| { readonly kind: "eqArray"; readonly rows: readonly OmmlContent[] }
	// m:acc.
	| { readonly kind: "accent"; readonly character: string; readonly base: OmmlContent }
	// m:bar.
	| { readonly kind: "bar"; readonly position: "top" | "bot"; readonly base: OmmlContent }
	// m:groupChr.
	| { readonly kind: "groupChar"; readonly character: string; readonly position: "top" | "bot"; readonly base: OmmlContent }
	// m:func.
	| { readonly kind: "function"; readonly name: OmmlContent; readonly base: OmmlContent }
	// m:borderBox, a box all round (Word's \rect).
	| { readonly kind: "borderBox"; readonly base: OmmlContent }
	// m:limLow or m:limUpp.
	| { readonly kind: "limit"; readonly position: "lower" | "upper"; readonly base: OmmlContent; readonly limit: OmmlContent }
	// At the top of a display equation: the next equation of its m:oMathPara (a line of its own). Nowhere else.
	| { readonly kind: "break" }

export interface OmmlWriteOptions {
	// m:oMathPara around the m:oMath (an equation on its own line).
	readonly display: boolean
}

// Objects nested in objects beyond this are not written (the LaTeX reader refuses deeper input).
export const MAX_OMML_NESTING = 32
// Elements below an m:oMath that readers read in full (core/math/omml.ts reads 96 argument levels).
export const MAX_OMML_DEPTH = 200

export const OMML_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math"

class Writer {
	out = ""
	private depth = 0

	empty(local: string, value?: string): void {
		this.out += value === undefined ? `<m:${local}/>` : `<m:${local} m:val="${escapeAttribute(value)}"/>`
	}

	// An argument element (m:e, m:num, m:sub...), empty when it has no content.
	argument(local: string, content: OmmlContent): void {
		if (content.length === 0) {
			this.out += `<m:${local}/>`

			return
		}

		this.out += `<m:${local}>`
		this.content(content)
		this.out += `</m:${local}>`
	}

	// Property element `local`, left out without properties.
	properties(local: string, write: () => void): void {
		const start = this.out.length

		this.out += `<m:${local}>`

		const inner = this.out.length

		write()

		if (this.out.length === inner) this.out = this.out.slice(0, start)
		else this.out += `</m:${local}>`
	}

	content(content: OmmlContent): void {
		for (const node of content) this.node(node)
	}

	node(node: OmmlNode): void {
		if (node.kind === "run") {
			this.run(node)

			return
		}

		if (node.kind === "break") return
		if (this.depth >= MAX_OMML_NESTING) throw new RangeError("An equation nested too deep")

		this.depth++
		this.object(node)
		this.depth--
	}

	run(node: Extract<OmmlNode, { kind: "run" }>): void {
		if (node.text === "") return

		this.out += "<m:r>"
		this.properties("rPr", () => {
			if (node.normal === true) {
				this.empty("nor")

				return
			}

			if (node.script !== undefined) this.empty("scr", node.script)
			if (node.style !== undefined) this.empty("sty", node.style)
		})

		const spaced = /^\s|\s$/.test(node.text)

		this.out += `<m:t${spaced ? ' xml:space="preserve"' : ""}>${escapeText(node.text)}</m:t></m:r>`
	}

	object(node: Exclude<OmmlNode, { kind: "run" | "break" }>): void {
		switch (node.kind) {
			case "fraction":
				this.out += "<m:f>"
				this.properties("fPr", () => {
					if (node.type !== undefined && node.type !== "bar") this.empty("type", node.type)
				})
				this.argument("num", node.numerator)
				this.argument("den", node.denominator)
				this.out += "</m:f>"
				return
			case "radical":
				this.out += "<m:rad>"
				this.properties("radPr", () => {
					if (node.degree === null) this.empty("degHide", "1")
				})
				this.argument("deg", node.degree ?? [])
				this.argument("e", node.base)
				this.out += "</m:rad>"
				return
			case "script":
				this.script(node)
				return
			case "nary":
				this.out += "<m:nary>"
				this.properties("naryPr", () => {
					this.empty("chr", node.operator)

					if (node.location !== null) this.empty("limLoc", node.location)
					if (node.sub === null) this.empty("subHide", "1")
					if (node.sup === null) this.empty("supHide", "1")
				})
				this.argument("sub", node.sub ?? [])
				this.argument("sup", node.sup ?? [])
				this.argument("e", node.base)
				this.out += "</m:nary>"
				return
			case "delimiter":
				this.out += "<m:d>"
				this.properties("dPr", () => {
					if (node.open !== "(") this.empty("begChr", node.open)
					if (node.separator !== undefined && node.separator !== "|") this.empty("sepChr", node.separator)
					if (node.close !== ")") this.empty("endChr", node.close)
				})

				for (const item of node.items.length === 0 ? [[]] : node.items) this.argument("e", item)

				this.out += "</m:d>"
				return
			case "matrix": {
				const columns = Math.max(1, ...node.rows.map(cells => cells.length))

				this.out += `<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="${String(columns)}"/><m:mcJc m:val="center"/></m:mcPr></m:mc></m:mcs></m:mPr>`

				for (const cells of node.rows.length === 0 ? [[]] : node.rows) {
					this.out += "<m:mr>"

					for (let k = 0; k < columns; k++) this.argument("e", cells[k] ?? [])

					this.out += "</m:mr>"
				}

				this.out += "</m:m>"
				return
			}
			case "eqArray":
				this.out += "<m:eqArr>"

				for (const row of node.rows.length === 0 ? [[]] : node.rows) this.argument("e", row)

				this.out += "</m:eqArr>"
				return
			case "accent":
				this.out += "<m:acc>"
				this.properties("accPr", () => this.empty("chr", node.character))
				this.argument("e", node.base)
				this.out += "</m:acc>"
				return
			case "bar":
				this.out += "<m:bar>"
				this.properties("barPr", () => {
					if (node.position === "top") this.empty("pos", "top")
				})
				this.argument("e", node.base)
				this.out += "</m:bar>"
				return
			case "groupChar":
				this.out += "<m:groupChr>"
				this.properties("groupChrPr", () => {
					this.empty("chr", node.character)

					if (node.position === "top") this.empty("pos", "top")
				})
				this.argument("e", node.base)
				this.out += "</m:groupChr>"
				return
			case "borderBox":
				this.out += "<m:borderBox>"
				this.argument("e", node.base)
				this.out += "</m:borderBox>"
				return
			case "function":
				this.out += "<m:func>"
				this.argument("fName", node.name)
				this.argument("e", node.base)
				this.out += "</m:func>"
				return
			case "limit": {
				const local = node.position === "lower" ? "limLow" : "limUpp"

				this.out += `<m:${local}>`
				this.argument("e", node.base)
				this.argument("lim", node.limit)
				this.out += `</m:${local}>`
				return
			}
		}
	}

	script(node: Extract<OmmlNode, { kind: "script" }>): void {
		if (node.pre === true) {
			this.out += "<m:sPre>"
			this.argument("sub", node.sub ?? [])
			this.argument("sup", node.sup ?? [])
			this.argument("e", node.base)
			this.out += "</m:sPre>"

			return
		}

		const local = node.sub === null ? (node.sup === null ? undefined : "sSup") : node.sup === null ? "sSub" : "sSubSup"

		if (local === undefined) {
			this.content(node.base)

			return
		}

		this.out += `<m:${local}>`
		this.argument("e", node.base)

		if (node.sub !== null) this.argument("sub", node.sub)
		if (node.sup !== null) this.argument("sup", node.sup)

		this.out += `</m:${local}>`
	}
}

// The equation's m:oMath (or m:oMathPara) element, m: prefixed and without namespace declarations. In a
// display equation each top-level break starts the next m:oMath. Throws a RangeError for content nested
// beyond MAX_OMML_NESTING.
export function writeOmml(content: OmmlContent, options: OmmlWriteOptions): string {
	const writer = new Writer()
	const lines: OmmlNode[][] = [[]]

	for (const node of content) {
		if (node.kind === "break") {
			if (options.display) lines.push([])
		} else {
			lines[lines.length - 1]?.push(node)
		}
	}

	for (const line of options.display ? lines : [lines.flat()]) {
		writer.out += "<m:oMath>"
		writer.content(line)
		writer.out += "</m:oMath>"
	}

	return options.display ? `<m:oMathPara>${writer.out}</m:oMathPara>` : writer.out
}
