import type { Mark } from "prosemirror-model"
import type { XmlDocument } from "../../xml/index"
import { EMPTY_RUN, type ColorProps, type PropsStore, type RunProps } from "../props/index"
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

	if (props.change !== undefined) {
		push(m.format_change.create({ id: props.change.id ?? null, author: props.change.author ?? null, date: props.change.date ?? null }))
	}

	return out
}
