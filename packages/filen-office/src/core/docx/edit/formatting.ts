import type { Mark } from "prosemirror-model"
import { decodeUtf8, type XmlDocument } from "../../xml/index"
import { partReader } from "../model/reader"
import {
	EMPTY_RUN,
	type ColorProps,
	type FontsProps,
	type LanguageProps,
	type PropsStore,
	type RunProps,
	type ThemeColorName
} from "../props/index"
import { docxSchema, type ColorAttrs } from "./schema"

// The source run formatting of one session. Each distinct w:rPr (by its bytes: the props store interns
// them) gets an id; text formatted with it carries a `run` mark naming its w:r and that id, plus the marks
// of the base's editable direct properties, which thousands of runs share. An edited run's w:rPr is
// rebuilt from the base's bytes with only the changed properties spliced in.
export interface RunBase {
	readonly id: number
	readonly props: RunProps
	// Where one copy of the w:rPr bytes is (every copy is byte-identical).
	readonly document: XmlDocument
	readonly node: number
	// The derived property marks, in schema order.
	readonly marks: readonly Mark[]
}

export class RunBases {
	private readonly byProps = new Map<RunProps, RunBase>()
	private readonly list: RunBase[] = []
	private readonly store: PropsStore

	constructor(store: PropsStore) {
		this.store = store
	}

	// The base for a w:rPr node, or undefined for none or an empty one (nothing to keep).
	of(document: XmlDocument, rPr: number): RunBase | undefined {
		if (rPr < 0) return undefined

		const props = this.store.run(document, rPr)

		if (props === EMPTY_RUN) return undefined

		let base = this.byProps.get(props)

		if (base === undefined) {
			const id = this.list.length

			base = { id, props, document, node: rPr, marks: formattingMarks(props) }
			this.byProps.set(props, base)
			this.list.push(base)
		}

		return base
	}

	get(id: number): RunBase | undefined {
		return this.list[id]
	}

	get size(): number {
		return this.list.length
	}
}

function colorAttrs(color: ColorProps | undefined): ColorAttrs | null {
	if (color === undefined) return null

	return {
		value: color.val ?? null,
		themeColor: color.themeColor ?? null,
		themeTint: color.themeTint ?? null,
		themeShade: color.themeShade ?? null
	}
}

function toggle(value: boolean | undefined, cs: boolean | undefined): { value: boolean | null; cs: boolean | null } | undefined {
	return value === undefined && cs === undefined ? undefined : { value: value ?? null, cs: cs ?? null }
}

// The marks of the direct properties a command can change, in schema order. A writer compares the marks
// a node has with these to find what to splice: equal marks mean an untouched property.
export function formattingMarks(props: RunProps): Mark[] {
	const m = docxSchema.marks
	const out: Mark[] = []
	const push = (mark: Mark): void => {
		out.push(mark)
	}

	if (props.styleId !== undefined) push(m.char_style.create({ styleId: props.styleId }))

	const bold = toggle(props.bold, props.boldCs)
	const italic = toggle(props.italic, props.italicCs)

	if (bold !== undefined) push(m.bold.create(bold))
	if (italic !== undefined) push(m.italic.create(italic))
	if (props.underline !== undefined) push(m.underline.create({ style: props.underline.style, ...colorAttrs(props.underline) }))
	if (props.strike !== undefined) push(m.strike.create({ value: props.strike }))
	if (props.doubleStrike !== undefined) push(m.double_strike.create({ value: props.doubleStrike }))
	if (props.caps !== undefined) push(m.caps.create({ value: props.caps }))
	if (props.smallCaps !== undefined) push(m.small_caps.create({ value: props.smallCaps }))
	if (props.vertAlign !== undefined) push(m.vert_align.create({ value: props.vertAlign }))
	if (props.color !== undefined) push(m.color.create(colorAttrs(props.color)))
	if (props.highlight !== undefined) push(m.highlight.create({ value: props.highlight }))

	if (props.shading !== undefined) {
		push(
			m.shading.create({
				pattern: props.shading.pattern,
				color: colorAttrs(props.shading.color),
				fill: colorAttrs(props.shading.fill)
			})
		)
	}

	if (props.fonts !== undefined) {
		const f = props.fonts

		push(
			m.font.create({
				ascii: f.ascii ?? null,
				hAnsi: f.hAnsi ?? null,
				eastAsia: f.eastAsia ?? null,
				cs: f.cs ?? null,
				asciiTheme: f.asciiTheme ?? null,
				hAnsiTheme: f.hAnsiTheme ?? null,
				eastAsiaTheme: f.eastAsiaTheme ?? null,
				csTheme: f.csTheme ?? null,
				hint: f.hint ?? null
			})
		)
	}

	if (props.size !== undefined || props.sizeCs !== undefined) push(m.size.create({ value: props.size ?? null, cs: props.sizeCs ?? null }))
	if (props.spacing !== undefined) push(m.spacing.create({ value: props.spacing }))

	if (props.lang !== undefined) {
		push(m.lang.create({ latin: props.lang.latin ?? null, eastAsia: props.lang.eastAsia ?? null, bidi: props.lang.bidi ?? null }))
	}

	if (props.noProof !== undefined) push(m.no_proof.create({ value: props.noProof }))
	if (props.webHidden !== undefined) push(m.web_hidden.create({ value: props.webHidden }))

	if (props.change !== undefined) {
		push(m.format_change.create({ id: props.change.id ?? null, author: props.change.author ?? null, date: props.change.date ?? null }))
	}

	return out
}

function colorProps(attrs: ColorAttrs | null): ColorProps | undefined {
	if (attrs === null) return undefined

	return {
		...(attrs.value !== null ? { val: attrs.value } : {}),
		...(attrs.themeColor !== null ? { themeColor: attrs.themeColor as ThemeColorName } : {}),
		...(attrs.themeTint !== null ? { themeTint: attrs.themeTint } : {}),
		...(attrs.themeShade !== null ? { themeShade: attrs.themeShade } : {})
	}
}

function defined<T extends object>(value: { [K in keyof T]: T[K] | null }): T {
	const out: Record<string, unknown> = {}

	for (const [key, entry] of Object.entries(value)) if (entry !== null) out[key] = entry

	return out as T
}

// The direct run properties the formatting marks of a node stand for (formattingMarks backwards). Properties
// without a mark (kerning, position, effects) are not among them.
export function runPropsOf(marksOf: readonly Mark[]): RunProps {
	const m = docxSchema.marks
	const out: Record<string, unknown> = {}

	for (const mark of marksOf) {
		const a = mark.attrs

		switch (mark.type) {
			case m.char_style:
				out["styleId"] = a["styleId"]

				break
			case m.bold:
			case m.italic: {
				const name = mark.type === m.bold ? "bold" : "italic"

				if (a["value"] !== null) out[name] = a["value"]
				if (a["cs"] !== null) out[`${name}Cs`] = a["cs"]

				break
			}
			case m.underline:
				out["underline"] = { style: a["style"], ...colorProps(a as unknown as ColorAttrs) }

				break
			case m.strike:
				out["strike"] = a["value"]

				break
			case m.double_strike:
				out["doubleStrike"] = a["value"]

				break
			case m.caps:
				out["caps"] = a["value"]

				break
			case m.small_caps:
				out["smallCaps"] = a["value"]

				break
			case m.vert_align:
				out["vertAlign"] = a["value"]

				break
			case m.color:
				out["color"] = colorProps(a as unknown as ColorAttrs)

				break
			case m.highlight:
				out["highlight"] = a["value"]

				break
			case m.shading:
				out["shading"] = {
					pattern: a["pattern"],
					...(a["color"] !== null ? { color: colorProps(a["color"] as ColorAttrs) } : {}),
					...(a["fill"] !== null ? { fill: colorProps(a["fill"] as ColorAttrs) } : {})
				}

				break
			case m.font:
				out["fonts"] = defined<FontsProps>(a as unknown as { [K in keyof FontsProps]: FontsProps[K] | null })

				break
			case m.size:
				if (a["value"] !== null) out["size"] = a["value"]
				if (a["cs"] !== null) out["sizeCs"] = a["cs"]

				break
			case m.spacing:
				out["spacing"] = a["value"]

				break
			case m.lang:
				out["lang"] = defined<LanguageProps>({
					latin: a["latin"] as string | null,
					eastAsia: a["eastAsia"] as string | null,
					bidi: a["bidi"] as string | null
				})

				break
			case m.no_proof:
				out["noProof"] = a["value"]

				break
			case m.web_hidden:
				out["webHidden"] = a["value"]

				break
		}
	}

	return out as RunProps
}

// Whether the w:rPr bytes of `base` keep their meaning written into `document`: the writer copies them as they
// are, so every prefix in them must be bound there (on its root, where Word declares them) to the same
// namespace.
export function baseFits(base: RunBase, document: XmlDocument): boolean {
	if (base.document === document) return true

	const from = partReader(base.document).namespaces
	const to = partReader(document).namespaces
	const root = document.documentElement()
	const source = base.document
	const fits = (node: number, prefix: string): boolean =>
		prefix === "xml" || (prefix !== "" && from.uri(node, prefix) === to.uri(root, prefix))
	const visit = (node: number): boolean => {
		if (!fits(node, source.prefix(node))) return false

		const spans = source.attributeSpans(node)

		for (let k = 0; k < spans.length; k += 4) {
			const name = decodeUtf8(source.bytes, spans[k] ?? 0, spans[k + 1] ?? 0)
			const colon = name.indexOf(":")

			// Unprefixed attributes are in no namespace; a declaration in the bytes is not carried by scope.
			if (name === "xmlns" || name.startsWith("xmlns:")) return false
			if (colon >= 0 && !fits(node, name.slice(0, colon))) return false
		}

		for (let child = source.firstChildElement(node); child >= 0; child = source.nextSiblingElement(child))
			if (!visit(child)) return false

		return true
	}

	try {
		return visit(base.node)
	} catch {
		return false
	}
}
