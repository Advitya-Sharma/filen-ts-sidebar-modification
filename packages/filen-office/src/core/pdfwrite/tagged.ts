import { PdfWriteError } from "./errors"
import { pdfName, type PdfDict, type PdfRef, type PdfValue } from "./objects"

// The logical structure of a tagged PDF (ISO 32000-1 §14.7-14.8; ISO 14289-1 for PDF/UA): a tree of
// structure elements in reading order whose leaves are marked-content sequences (MCIDs on a page) and
// annotations (OBJR). Painters open and close elements in document order; an element may span pages.
//
// The standard types of §14.8.4 are used as is; any other type needs `role`, a standard type it maps to
// (written to the RoleMap). Headings are H1-H6 and must not skip levels (ISO 14289-1 §7.4.2); figures and
// formulas need `alt`; table header cells need `scope`.

export const STANDARD_TYPES = new Set([
	"Document",
	"Part",
	"Art",
	"Sect",
	"Div",
	"BlockQuote",
	"Caption",
	"TOC",
	"TOCI",
	"Index",
	"NonStruct",
	"Private",
	"P",
	"H",
	"H1",
	"H2",
	"H3",
	"H4",
	"H5",
	"H6",
	"L",
	"LI",
	"Lbl",
	"LBody",
	"Table",
	"TR",
	"TH",
	"TD",
	"THead",
	"TBody",
	"TFoot",
	"Span",
	"Quote",
	"Note",
	"Reference",
	"BibEntry",
	"Code",
	"Link",
	"Annot",
	"Ruby",
	"RB",
	"RT",
	"RP",
	"Warichu",
	"WT",
	"WP",
	"Figure",
	"Formula",
	"Form"
])

export interface PdfStructOptions {
	// A standard type the element's own type maps to (required for a non-standard type).
	readonly role?: string
	// Alternate description (§14.9.3): required for Figure and Formula.
	readonly alt?: string
	// Replacement text (§14.9.4).
	readonly actualText?: string
	// Expansion of an abbreviation (§14.9.5).
	readonly expansion?: string
	// BCP 47 language when it differs from the document's.
	readonly lang?: string
	readonly title?: string
	// Table attributes (§14.8.5.7).
	readonly scope?: "Row" | "Column" | "Both"
	readonly rowSpan?: number
	readonly colSpan?: number
	// List numbering (§14.8.5.5) for an L element.
	readonly listNumbering?: "None" | "Disc" | "Circle" | "Square" | "Decimal" | "UpperRoman" | "LowerRoman" | "UpperAlpha" | "LowerAlpha"
	// Layout bounding box (px of the page it first appears on), for Figure, Formula and Table.
	readonly bbox?: readonly [number, number, number, number]
}

// A kid of an element: another element (its index, >= 0), marked content on a page (-1 - (page * MCID_SPAN +
// mcid)), or an annotation. Numbers, not objects: a tagged spreadsheet holds a structure element per cell.
type Kid = number | { readonly kind: "objr"; readonly page: number; readonly ref: PdfRef }

const MCID_SPAN = 2 ** 24

function mcrKid(page: number, mcid: number): number {
	return -1 - (page * MCID_SPAN + mcid)
}

const NO_OPTIONS: PdfStructOptions = Object.freeze({})

interface Node {
	readonly type: string
	readonly parent: number
	readonly options: PdfStructOptions
	kids: Kid[]
	page: number
	closed: boolean
}

export class StructTree {
	private readonly nodes: Node[] = []
	private readonly stack: number[] = []
	private readonly roles = new Map<string, string>()
	// Per page index: its parent tree key, and the element of each MCID.
	private readonly pageKeys = new Map<number, number>()
	private readonly pageMcids = new Map<number, number[]>()
	private readonly annotationKeys: { key: number; node: number }[] = []
	private nextKey = 0

	constructor() {
		this.nodes.push({ type: "Document", parent: -1, options: {}, kids: [], page: -1, closed: false })
		this.stack.push(0)
	}

	get root(): number {
		return 0
	}

	get current(): number {
		return this.stack[this.stack.length - 1] ?? 0
	}

	typeOf(node: number): string {
		return this.nodes[node]?.type ?? "Span"
	}

	open(type: string, options: PdfStructOptions = NO_OPTIONS): number {
		if (!STANDARD_TYPES.has(type)) {
			const role = options.role

			if (role === undefined || !STANDARD_TYPES.has(role))
				throw new PdfWriteError("state", `Structure type ${type} needs a standard role`)

			const known = this.roles.get(type)

			if (known !== undefined && known !== role)
				throw new PdfWriteError("state", `Structure type ${type} is already mapped to ${known}`)

			this.roles.set(type, role)
		}

		const parent = this.current
		const id = this.nodes.length

		this.nodes.push({ type, parent, options, kids: [], page: -1, closed: false })
		this.nodes[parent]?.kids.push(id)
		this.stack.push(id)

		return id
	}

	close(node: number): void {
		if (this.current !== node || node === 0)
			throw new PdfWriteError("unbalanced", "Structure elements must close in the order they opened")

		this.stack.pop()

		const entry = this.nodes[node]

		if (entry !== undefined) {
			entry.closed = true
			// Kids arrays grow in steps; a closed element keeps exactly its own.
			entry.kids = entry.kids.slice()
		}
	}

	// A new MCID on `page` for content of `node`.
	mark(node: number, page: number): number {
		const entry = this.nodes[node]

		if (entry === undefined || entry.closed)
			throw new PdfWriteError("state", "Marked content for a closed or unknown structure element")

		let mcids = this.pageMcids.get(page)

		if (mcids === undefined) {
			mcids = []
			this.pageMcids.set(page, mcids)
			this.pageKeys.set(page, this.nextKey++)
		}

		const mcid = mcids.length

		mcids.push(node)
		entry.kids.push(mcrKid(page, mcid))

		if (entry.page === -1) entry.page = page

		return mcid
	}

	pageKey(page: number): number | undefined {
		return this.pageKeys.get(page)
	}

	// The /StructParent key of an annotation on `page` belonging to `node`.
	annotation(node: number, page: number, ref: PdfRef): number {
		const entry = this.nodes[node]

		if (entry === undefined) throw new PdfWriteError("state", "Annotation for an unknown structure element")

		entry.kids.push({ kind: "objr", page, ref })

		if (entry.page === -1) entry.page = page

		const key = this.nextKey++

		this.annotationKeys.push({ key, node })

		return key
	}

	// Writes the elements, the parent tree and the root; returns the StructTreeRoot ref. `pause` runs every
	// few hundred elements (object streams flushed, the host given a turn).
	async write(
		allocate: () => PdfRef,
		object: (ref: PdfRef, value: PdfValue) => void,
		pause: () => Promise<void>,
		pages: readonly PdfRef[],
		pageHeights: readonly number[]
	): Promise<PdfRef> {
		if (this.stack.length !== 1) throw new PdfWriteError("unbalanced", "Structure elements left open")

		const rootRef = allocate()
		const refs = this.nodes.map(() => allocate())
		// Pages of descendants, for elements whose own content starts deeper.
		const firstPage = (node: number): number => {
			const entry = this.nodes[node]

			if (entry === undefined) return -1
			if (entry.page !== -1) return entry.page

			for (const kid of entry.kids) {
				const page = typeof kid !== "number" ? kid.page : kid >= 0 ? firstPage(kid) : Math.floor((-1 - kid) / MCID_SPAN)

				if (page !== -1) return page
			}

			return -1
		}

		for (let index = 0; index < this.nodes.length; index++) {
			const node = this.nodes[index]
			const ref = refs[index]

			if (node === undefined || ref === undefined) continue
			if (index % 256 === 255) await pause()

			const page = firstPage(index)
			const pageRef = page === -1 ? undefined : pages[page]
			const kids: PdfValue[] = node.kids.map(kid => {
				if (typeof kid === "number") {
					if (kid >= 0) return refs[kid] ?? null

					const kidPage = Math.floor((-1 - kid) / MCID_SPAN)
					const mcid = (-1 - kid) % MCID_SPAN

					return kidPage === page ? mcid : { Type: pdfName("MCR"), Pg: pages[kidPage] ?? null, MCID: mcid }
				}

				return { Type: pdfName("OBJR"), Pg: pages[kid.page] ?? null, Obj: kid.ref }
			})
			const o = node.options
			const attributes: PdfDict[] = []

			if (o.scope !== undefined || o.rowSpan !== undefined || o.colSpan !== undefined) {
				attributes.push({
					O: pdfName("Table"),
					Scope: o.scope === undefined ? undefined : pdfName(o.scope),
					RowSpan: o.rowSpan,
					ColSpan: o.colSpan
				})
			}
			if (o.listNumbering !== undefined) attributes.push({ O: pdfName("List"), ListNumbering: pdfName(o.listNumbering) })
			if (o.bbox !== undefined && page !== -1) {
				const height = pageHeights[page] ?? 0
				const [x0, y0, x1, y1] = o.bbox

				attributes.push({ O: pdfName("Layout"), BBox: [x0 * 0.75, height - y1 * 0.75, x1 * 0.75, height - y0 * 0.75] })
			}

			object(ref, {
				Type: pdfName("StructElem"),
				S: pdfName(node.type),
				P: index === 0 ? rootRef : (refs[node.parent] ?? rootRef),
				Pg: pageRef,
				K: kids.length === 1 ? kids[0] : kids,
				Alt: o.alt,
				ActualText: o.actualText,
				E: o.expansion,
				Lang: o.lang,
				T: o.title,
				A: attributes.length === 0 ? undefined : attributes.length === 1 ? attributes[0] : attributes
			})
		}

		const nums: PdfValue[] = []
		const entries: { key: number; value: PdfValue }[] = []

		for (const [page, mcids] of this.pageMcids) {
			const key = this.pageKeys.get(page)

			if (key !== undefined) entries.push({ key, value: mcids.map(node => refs[node] ?? null) })
		}

		for (const { key, node } of this.annotationKeys) entries.push({ key, value: refs[node] ?? null })

		entries.sort((a, b) => a.key - b.key)

		for (const entry of entries) nums.push(entry.key, entry.value)

		const parentTree = allocate()

		object(parentTree, { Nums: nums })

		const roleMap: Record<string, PdfValue> = {}

		for (const [type, role] of this.roles) roleMap[type] = pdfName(role)

		object(rootRef, {
			Type: pdfName("StructTreeRoot"),
			K: refs[0] ?? null,
			ParentTree: parentTree,
			ParentTreeNextKey: this.nextKey,
			RoleMap: this.roles.size === 0 ? undefined : roleMap
		})

		return rootRef
	}
}
