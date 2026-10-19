import { bytesSource, OpcPackage } from "@filen/office/core/opc"
import { Namespaces, NodeKind, parseXml, type XmlDocument } from "@filen/office/core/xml"

// The conformance catalogue: constructs a writer puts into a WordprocessingML package that Word refuses, repairs
// or reads differently from the standard, from [MS-OI29500] notes, [MS-DOCX] and ECMA-376 constraints the
// schemas cannot express (each item names its source). checkConformance names every violation
// in a package, counted, so a save is compared with its source: a save may not have more of any.

export interface CatalogueItem {
	readonly id: string
	readonly title: string
	// ECMA-376 / ISO/IEC 29500 clause and the Microsoft note (MS-OI29500 or MS-DOCX section and letter).
	readonly source: string
}

export const CATALOGUE: readonly CatalogueItem[] = [
	{ id: "C01", title: "annotation dates in Word's form", source: "Part 1 §17.13.4.2; MS-OI29500 2.1.313 d; MS-DOCX 2.10.3.1, 2.12.2.1" },
	{ id: "C02", title: "ST_OnOff written as 1/0/true/false, never on/off", source: "Part 1 §22.9.2.7; MS-OI29500 2.1.1744 a" },
	{ id: "C03", title: "jc written as left/right, never start/end", source: "Part 1 §17.3.1.13, §17.4.28, §17.18.44-45" },
	{
		id: "C04",
		title: "measures as integers: no universal measures, pct in fiftieths",
		source: "Part 1 §17.4.63, §17.4.71, §22.9.2.14; MS-OI29500 2.1.166, 2.1.171"
	},
	{ id: "C05", title: "tblLook val agrees with its attributes", source: "Part 1 §17.4.55; MS-OI29500 2.1.160 c" },
	{
		id: "C06",
		title: "notes parts: separators, id range, resolvable references",
		source: "Part 1 §17.11; MS-OI29500 2.1.300 a, 2.1.301 a-b"
	},
	{
		id: "C07",
		title: "numbering: abstractNum before num, ids resolve, abstractNumId not negative, nsid unique",
		source: "Part 1 §17.9.1, §17.9.15-16; MS-OI29500 2.1.275 a, 2.1.285 a"
	},
	{
		id: "C08",
		title: "wp:docPr ids unique across stories and within 1..2147483646",
		source: "Part 1 §20.4.2.5; MS-OI29500 2.1.1351 a-b"
	},
	{
		id: "C09",
		title: "annotation ids: unique per kind, ranges paired, no -1, at most 32767 comment ranges",
		source: "Part 1 §17.13; MS-OI29500 2.1.317 a, 2.1.315 a, 2.1.358 c, 2.1.340 a"
	},
	{ id: "C10", title: "field characters balanced per story", source: "Part 1 §17.16.18" },
	{
		id: "C11",
		title: "tables: rows, tblGrid, cells end in a paragraph, at most 63 cells, rows within the grid",
		source: "Part 1 §17.4.37, §17.4.48, §17.4.65; MS-OI29500 2.1.149, 2.1.168 b"
	},
	{ id: "C12", title: "sectPr only where Word allows it", source: "Part 1 §17.6.17-18; MS-OI29500 2.1.224 c" },
	{
		id: "C13",
		title: "hyperlinks: r:id resolves to an External hyperlink relationship; attribute lengths",
		source: "Part 1 §17.16.22; MS-OI29500 2.1.521 a"
	},
	{ id: "C14", title: "bookmark names at most 40 characters and unique", source: "Part 1 §17.13.6.2; MS-OI29500 2.1.356 b-c" },
	{ id: "C15", title: "content control ids unique 32-bit integers", source: "Part 1 §17.5.2.18; MS-OI29500 2.1.196 a" },
	{ id: "C16", title: "w14:paraId/textId in 1..0x7FFFFFFF and unique per part", source: "MS-DOCX 2.6.2.3, 2.6.2.4" },
	{
		id: "C17",
		title: "comment extension parts linked: commentsExtended, commentsIds, commentsExtensible, people",
		source: "MS-DOCX 2.5.3.1, 2.5.3.5, 2.8.3.1, 2.10.3.1"
	},
	{
		id: "C18",
		title: "documentProtection hash attributes complete, spin count in range",
		source: "Part 1 §17.15.1.29; MS-OI29500 2.1.402 d"
	},
	{
		id: "C19",
		title: "Markup Compatibility prefixes declared; none in properties parts",
		source: "Part 3 §7, §10; MS-OI29500 2.1.1753 a"
	},
	{
		id: "C20",
		title: "only XML 1.0 characters, no raw carriage returns",
		source: "XML 1.0 §2.2; Part 1 §22.9.2.19; MS-OI29500 2.1.1747 b"
	},
	{ id: "C21", title: "every part has one content type", source: "Part 2 §10.1.2" },
	{
		id: "C22",
		title: "relationships: unique ids, internal targets exist, references resolve",
		source: "Part 2 §9.3; Part 1 §17.16.22, §20.4.2"
	},
	{
		id: "C23",
		title: "comment initials at most 9 characters, authors at most 255",
		source: "Part 1 §17.13.4.2; MS-OI29500 2.1.313 a, c"
	},
	{ id: "C24", title: "tracked changes hold runs only", source: "Part 1 §17.13.5.14, §17.13.5.18 (CT_RunTrackChange)" },
	{ id: "C25", title: "line numbering: countBy 1..100, start 0..32767", source: "Part 1 §17.6.8; MS-OI29500 Part 1 §17.6.8 notes a, c" }
]

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
const W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
// Word 2013's namespace, and the one its prerelease wrote (files of that time carry it).
const W15_PRERELEASE = "http://schemas.microsoft.com/office/word/2010/11/wordml"
const W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
const W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
const W16DU = "http://schemas.microsoft.com/office/word/2023/wordml/word16du"
const PACKAGE_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
const HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml."
const MACRO_MAIN = /^application\/vnd\.ms-word\.(document|template)\.macroEnabled(Template)?\.main\+xml$/
const RELS_TYPE = "application/vnd.openxmlformats-package.relationships+xml"
const PROPERTIES = new Set([
	"application/vnd.openxmlformats-package.core-properties+xml",
	"application/vnd.openxmlformats-officedocument.extended-properties+xml",
	"application/vnd.openxmlformats-officedocument.custom-properties+xml"
])
const EXTENSION_RELATIONSHIPS: Readonly<Record<string, string>> = {
	commentsExtended: "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
	commentsIds: "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
	commentsExtensible: "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
	people: "http://schemas.microsoft.com/office/2011/relationships/people"
}

const REVISIONS = new Set([
	"ins",
	"del",
	"moveFrom",
	"moveTo",
	"rPrChange",
	"pPrChange",
	"sectPrChange",
	"tblPrChange",
	"tblPrExChange",
	"trPrChange",
	"tcPrChange",
	"tblGridChange",
	"numberingChange",
	"cellIns",
	"cellDel",
	"cellMerge"
])
const DATED = new Set([...REVISIONS, "comment", "moveFromRangeStart", "moveToRangeStart"])
// What CT_RunTrackChange holds (EG_ContentRunContent, EG_RunLevelElements; math is the m namespace's).
const RUN_TRACKED = new Set([
	"r",
	"customXml",
	"smartTag",
	"sdt",
	"dir",
	"bdo",
	"proofErr",
	"permStart",
	"permEnd",
	"bookmarkStart",
	"bookmarkEnd",
	"moveFromRangeStart",
	"moveFromRangeEnd",
	"moveToRangeStart",
	"moveToRangeEnd",
	"commentRangeStart",
	"commentRangeEnd",
	"customXmlInsRangeStart",
	"customXmlInsRangeEnd",
	"customXmlDelRangeStart",
	"customXmlDelRangeEnd",
	"customXmlMoveFromRangeStart",
	"customXmlMoveFromRangeEnd",
	"customXmlMoveToRangeStart",
	"customXmlMoveToRangeEnd",
	"ins",
	"del",
	"moveFrom",
	"moveTo"
])
// Elements of type CT_TblWidth (w:w with w:type).
const WIDTHS = new Set(["tblW", "tcW", "tblInd", "tblCellSpacing", "wBefore", "wAfter", "top", "left", "bottom", "right", "start", "end"])
// Block-level content a cell may end with; range markup around it does not count.
const BLOCKS = new Set(["p", "tbl", "sdt", "customXml", "altChunk"])
const UNIVERSAL_MEASURE = /^-?\d+(\.\d+)?(mm|cm|in|pt|pc|pi)$/
const WORD_DATE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/
const LOOK_BITS: readonly (readonly [string, number])[] = [
	["firstRow", 0x20],
	["lastRow", 0x40],
	["firstColumn", 0x80],
	["lastColumn", 0x100],
	["noHBand", 0x200],
	["noVBand", 0x400]
]

export interface ConformanceReport {
	// "<item> <what>" -> how often.
	readonly violations: Map<string, number>
	// Item -> constructs it checked in the parts `changed` names (all parts without it).
	readonly checked: Map<string, number>
}

class Report implements ConformanceReport {
	readonly violations = new Map<string, number>()
	readonly checked = new Map<string, number>()

	add(key: string, count = 1): void {
		if (count > 0) this.violations.set(key, (this.violations.get(key) ?? 0) + count)
	}

	check(item: string, counted: boolean): void {
		if (counted) this.checked.set(item, (this.checked.get(item) ?? 0) + 1)
	}

	// Counts every value seen more than once (one per extra occurrence).
	duplicates(key: string, values: Iterable<string>): void {
		const seen = new Set<string>()
		let extra = 0

		for (const value of values) {
			if (seen.has(value)) extra++
			else seen.add(value)
		}

		this.add(key, extra)
	}
}

interface Relationship {
	readonly type: string
	readonly external: boolean
	// The resolved part name of an internal target.
	readonly target: string
}

interface Part {
	readonly name: string
	readonly type: string
	readonly doc: XmlDocument
	readonly ns: Namespaces
	readonly changed: boolean
}

function partKey(name: string): string {
	return decodeURIComponent(name.replace(/^\//, "")).toLowerCase()
}

function resolveTarget(source: string, target: string): string {
	const base = target.startsWith("/") ? [] : source.split("/").slice(0, -1)

	for (const segment of target.replace(/^\//, "").split("/")) {
		if (segment === "..") base.pop()
		else if (segment !== "." && segment !== "") base.push(segment)
	}

	return base.join("/")
}

function relsSource(rels: string): string {
	const match = /^(.*?)_rels\/([^/]+)\.rels$/i.exec(rels)

	return match === null ? "" : `${match[1] ?? ""}${match[2] ?? ""}`
}

function* children(doc: XmlDocument, node: number): Generator<number> {
	for (let child = doc.firstChild(node); child >= 0; child = doc.nextSibling(child))
		if (doc.kinds[child] === NodeKind.Element) yield child
}

// Attributes of an element by "{uri}local" ("local" when unprefixed).
function attributesOf(part: Part, node: number): Map<string, string> {
	const out = new Map<string, string>()

	for (const { name, value } of part.doc.attributes(node)) {
		const colon = name.indexOf(":")

		if (colon < 0) out.set(name, value)
		else if (!name.startsWith("xmlns:")) out.set(`{${part.ns.uri(node, name.slice(0, colon)) ?? "?"}}${name.slice(colon + 1)}`, value)
	}

	return out
}

const w = (local: string): string => `{${W}}${local}`

function w15(attributes: Map<string, string>, local: string): string | undefined {
	return attributes.get(`{${W15}}${local}`) ?? attributes.get(`{${W15_PRERELEASE}}${local}`)
}

interface Story {
	// Open fields: whether each has had its separate.
	readonly fields: boolean[]
}

interface Collected {
	readonly docPrIds: string[]
	readonly bookmarkNames: string[]
	readonly sdtIds: string[]
	readonly commentRefs: string[]
	readonly commentRanges: Map<string, number>
	readonly footnoteRefs: string[]
	readonly endnoteRefs: string[]
	readonly numIdRefs: string[]
	readonly authors: Set<string>
	// Comment id -> paraId of its last paragraph.
	readonly commentParaIds: Map<string, string | undefined>
	readonly commentIds: string[]
	commentRangeCount: number
}

function storyEnd(report: Report, story: Story): void {
	report.add("C10 field begun but not ended in its story", story.fields.length)
	story.fields.length = 0
}

// Walks one WordprocessingML part in document order, Fallback branches left out (the Choice is what Word reads).
function walkPart(part: Part, report: Report, collected: Collected, relationships: Map<string, Relationship>): void {
	const { doc, ns, changed } = part
	const kind = part.type.startsWith(WML) ? part.type.slice(WML.length).replace(/\+xml$/, "") : part.type
	const storyPart =
		["header", "footer", "document.main", "template.main", "footnotes", "endnotes", "comments", "document.glossary"].includes(kind) ||
		/macroEnabled/.test(kind)
	const paraIds: string[] = []
	const revisionIds: string[] = []
	const bookmarkIds = new Map<string, number>()
	const moveRanges = new Map<string, number>()
	const permRanges = new Map<string, number>()
	const moveNames: { from: Set<string>; to: Set<string> } = { from: new Set(), to: new Set() }
	const root = doc.documentElement()
	let story: Story = { fields: [] }

	const bookmarkStarts: string[] = []
	// `fallback`: inside an mc:Fallback, whose content repeats the Choice's (ids included): only what each
	// element must satisfy on its own is checked there.
	const visit = (node: number, inCell: boolean, inTextbox: boolean, fallback: boolean): void => {
		const uri = ns.elementUri(node)
		const local = doc.localName(node)
		const attributes = attributesOf(part, node)

		// C22: relationship references resolve in the part's relationships.
		for (const [name, value] of attributes) {
			if (!name.startsWith(`{${R}}`)) continue

			report.check("C22", changed)

			if (value !== "" && !relationships.has(value)) report.add("C22 relationship reference unresolved")
		}

		// C19: prefixes MC attributes name are declared.
		for (const [name, value] of attributes) {
			if (!name.startsWith(`{${MC}}`) && !(uri === MC && local === "Choice" && name === "Requires")) continue

			report.check("C19", changed)

			for (const token of value.split(/\s+/)) {
				if (token === "") continue

				const prefix = token.includes(":") ? token.slice(0, token.indexOf(":")) : token

				if (ns.uri(node, prefix) === undefined) report.add(`C19 undeclared prefix in ${name.slice(name.indexOf("}") + 1)}`)
			}
		}

		if (fallback || (uri === MC && local === "Fallback")) {
			for (let child = doc.firstChild(node); child >= 0; child = doc.nextSibling(child)) {
				if (doc.kinds[child] === NodeKind.Element) visit(child, inCell, inTextbox, true)
			}

			return
		}

		if (attributes.has(`{${W14}}paraId`)) {
			report.check("C16", changed)
			paraIds.push(attributes.get(`{${W14}}paraId`) ?? "")
		} else if (attributes.has(`{${W14}}textId`) && local === "p") report.add("C16 textId without paraId")

		for (const id of ["paraId", "textId"]) {
			const value = attributes.get(`{${W14}}${id}`)

			if (value === undefined) continue

			const number = /^[0-9A-Fa-f]{1,8}$/.test(value) ? Number.parseInt(value, 16) : Number.NaN

			if (!(number > 0 && number < 0x80000000)) report.add(`C16 ${id} out of range`)
		}

		if (uri === W) {
			const val = attributes.get(w("val"))

			if (val === "on" || val === "off") report.add("C02 ST_OnOff written as on/off")
			if (val !== undefined && changed && /^(true|false|0|1)$/.test(val)) report.check("C02", true)

			for (const [name, value] of attributes) {
				if (name.startsWith(`{${W}}`) && UNIVERSAL_MEASURE.test(value)) report.add("C04 universal measure")
			}

			if (DATED.has(local) && attributes.has(w("date"))) {
				report.check("C01", changed)

				if (!WORD_DATE.test(attributes.get(w("date")) ?? "")) report.add("C01 annotation date not in Word's form")
			}

			const dateUtc = attributes.get(`{${W16DU}}dateUtc`)

			if (dateUtc !== undefined && !WORD_DATE.test(dateUtc)) report.add("C01 dateUtc not UTC")

			if ((local === "ins" || local === "del" || local === "moveFrom" || local === "moveTo") && !inProperties(doc, node)) {
				report.check("C24", changed)

				for (const child of children(doc, node)) {
					const name = doc.localName(child)

					if (ns.elementUri(child) === W && !RUN_TRACKED.has(name)) report.add(`C24 ${name} inside a tracked change`)
				}
			}

			if (REVISIONS.has(local) && attributes.has(w("id"))) {
				const id = attributes.get(w("id")) ?? ""

				revisionIds.push(id)

				if (id === "-1") report.add("C09 revision id -1")
				else if (!/^-?\d+$/.test(id)) report.add("C09 revision id not an integer")
			}

			if ((WIDTHS.has(local) && attributes.has(w("type"))) || local === "gridCol") report.check("C04", changed)

			if (WIDTHS.has(local) && attributes.get(w("type")) === "pct") {
				if ((attributes.get(w("w")) ?? "").endsWith("%")) report.add("C04 pct width as a percent string")
			}

			switch (local) {
				case "lnNumType": {
					report.check("C25", changed)

					const countBy = attributes.get(w("countBy"))
					const start = attributes.get(w("start"))

					if (countBy !== undefined && !(/^\d+$/.test(countBy) && Number(countBy) >= 1 && Number(countBy) <= 100))
						report.add("C25 lnNumType countBy outside 1..100")
					if (start !== undefined && !(/^\d+$/.test(start) && Number(start) <= 32767))
						report.add("C25 lnNumType start outside 0..32767")
					break
				}
				case "jc": {
					report.check("C03", changed)

					if (val === "start" || val === "end") report.add("C03 jc start/end")
					break
				}
				case "tblLook": {
					report.check("C05", changed)

					const present = LOOK_BITS.filter(([name]) => attributes.has(w(name)))

					if (val !== undefined && present.length > 0 && /^[0-9A-Fa-f]{1,4}$/.test(val)) {
						const bits = Number.parseInt(val, 16)

						for (const [name, bit] of present) {
							const on = /^(1|true|on)$/.test(attributes.get(w(name)) ?? "")

							if (on !== ((bits & bit) !== 0)) {
								report.add("C05 tblLook val and attributes disagree")
								break
							}
						}
					}
					break
				}
				case "bookmarkStart": {
					const id = attributes.get(w("id")) ?? ""
					const name = attributes.get(w("name")) ?? ""

					report.check("C14", changed)
					report.check("C09", changed)
					bookmarkStarts.push(id)
					bookmarkIds.set(id, (bookmarkIds.get(id) ?? 0) + 1)
					collected.bookmarkNames.push(name)

					if ([...name].length > 40) report.add("C14 bookmark name over 40 characters")
					break
				}
				case "bookmarkEnd": {
					const id = attributes.get(w("id")) ?? ""

					bookmarkIds.set(id, (bookmarkIds.get(id) ?? 0) - 1)
					break
				}
				case "moveFromRangeStart":
				case "moveToRangeStart":
				case "moveFromRangeEnd":
				case "moveToRangeEnd": {
					const id = `${local.startsWith("moveFrom") ? "from" : "to"} ${attributes.get(w("id")) ?? ""}`

					moveRanges.set(id, (moveRanges.get(id) ?? 0) + (local.endsWith("Start") ? 1 : -1))

					const name = attributes.get(w("name"))

					if (name !== undefined) moveNames[local.startsWith("moveFrom") ? "from" : "to"].add(name)
					break
				}
				case "permStart":
				case "permEnd": {
					const id = attributes.get(w("id")) ?? ""

					permRanges.set(id, (permRanges.get(id) ?? 0) + (local === "permStart" ? 1 : -1))

					if (!/^-?\d+$/.test(id) || Math.abs(Number(id)) > 2147483647) report.add("C09 permission id not a 32-bit integer")
					break
				}
				case "commentRangeStart":
				case "commentRangeEnd": {
					const id = attributes.get(w("id")) ?? ""

					report.check("C09", changed)
					collected.commentRanges.set(id, (collected.commentRanges.get(id) ?? 0) + (local === "commentRangeStart" ? 1 : -1))

					if (local === "commentRangeStart") collected.commentRangeCount++
					break
				}
				case "commentReference":
					collected.commentRefs.push(attributes.get(w("id")) ?? "")
					break
				case "footnoteReference":
					collected.footnoteRefs.push(attributes.get(w("id")) ?? "")
					break
				case "endnoteReference":
					collected.endnoteRefs.push(attributes.get(w("id")) ?? "")
					break
				case "numId": {
					const parent = doc.parent(node)

					if (parent >= 0 && doc.localName(parent) === "numPr") collected.numIdRefs.push(val ?? "")
					break
				}
				case "sdtPr": {
					report.check("C15", changed)

					for (const child of children(doc, node)) {
						if (ns.elementUri(child) !== W || doc.localName(child) !== "id") continue

						const id = attributesOf(part, child).get(w("val")) ?? ""

						collected.sdtIds.push(id)

						if (!/^-?\d+$/.test(id) || Number(id) > 2147483647 || Number(id) < -2147483648)
							report.add("C15 sdt id not a 32-bit integer")
					}
					break
				}
				case "fldChar": {
					report.check("C10", changed)

					const type = attributes.get(w("fldCharType"))

					if (type === "begin") story.fields.push(false)
					else if (type === "separate") {
						if (story.fields.length === 0) report.add("C10 field separate outside a field")
						else if (story.fields[story.fields.length - 1] === true) report.add("C10 field separated twice")
						else story.fields[story.fields.length - 1] = true
					} else if (type === "end") {
						if (story.fields.length === 0) report.add("C10 field end without begin")
						else story.fields.pop()
					}
					break
				}
				case "hyperlink": {
					report.check("C13", changed)

					const id = attributes.get(`{${R}}id`)

					if (id !== undefined) {
						const relationship = relationships.get(id)

						if (relationship === undefined || relationship.type !== HYPERLINK)
							report.add("C13 hyperlink r:id not a hyperlink relationship")
						else if (!relationship.external) report.add("C13 hyperlink relationship not External")
					}

					for (const [name, limit] of [
						["anchor", 255],
						["docLocation", 255],
						["tgtFrame", 255],
						["tooltip", 260]
					] as const) {
						if ([...(attributes.get(w(name)) ?? "")].length > limit)
							report.add(`C13 hyperlink ${name} over ${limit} characters`)
					}
					break
				}
				case "sectPr": {
					const parent = doc.parent(node)

					if (parent >= 0 && doc.localName(parent) === "pPr" && doc.localName(doc.parent(parent)) === "p") {
						report.check("C12", changed)

						if (inCell || inTextbox || (!["document.main", "template.main"].includes(kind) && !/macroEnabled/.test(kind)))
							report.add("C12 sectPr in a paragraph Word does not allow it in")
					} else if (parent >= 0 && doc.localName(parent) === "body") {
						report.check("C12", changed)

						let next = doc.nextSibling(node)

						while (next >= 0 && doc.kinds[next] !== NodeKind.Element) next = doc.nextSibling(next)

						if (next >= 0) report.add("C12 body sectPr not last")
					}
					break
				}
				case "tbl": {
					report.check("C11", changed)

					const grid = [...children(doc, node)].find(child => ns.elementUri(child) === W && doc.localName(child) === "tblGrid")
					const columns =
						grid === undefined ? 0 : [...children(doc, grid)].filter(child => doc.localName(child) === "gridCol").length

					if (grid === undefined) report.add("C11 table without tblGrid")

					const rows = rowsOf(part, node)

					if (rows.length === 0) report.add("C11 table without rows")

					for (const row of rows) {
						let span = 0
						let cells = 0

						for (const item of cellsOf(part, row)) {
							if (typeof item === "number") {
								span += item
								continue
							}

							cells++
							span += item.span
						}

						if (cells > 63) report.add("C11 row over 63 cells")
						if (grid !== undefined && span > columns) report.add("C11 row wider than its grid")
					}
					break
				}
				case "tc": {
					if (!endsInParagraph(part, node)) report.add("C11 cell not ending in a paragraph")
					break
				}
				case "documentProtection":
				case "writeProtection": {
					report.check("C18", changed)

					const spin = attributes.get(w("cryptSpinCount")) ?? attributes.get(w("spinCount"))

					if (spin !== undefined && !(/^\d+$/.test(spin) && Number(spin) <= 10_000_000)) report.add("C18 spin count out of range")

					const modern = ["algorithmName", "hashValue", "saltValue", "spinCount"].filter(name => attributes.has(w(name)))
					const legacy = ["cryptAlgorithmSid", "cryptSpinCount", "hash", "salt"].filter(name => attributes.has(w(name)))

					if ((modern.length > 0 && modern.length < 4) || (legacy.length > 0 && legacy.length < 4))
						report.add("C18 incomplete hash attributes")
					break
				}
				case "comment": {
					if (doc.parent(node) !== root) break

					report.check("C23", changed)

					const author = attributes.get(w("author"))
					const initials = attributes.get(w("initials"))

					if (author !== undefined) collected.authors.add(author)
					if (initials !== undefined && [...initials].length > 9) report.add("C23 comment initials over 9 characters")
					if (author !== undefined && [...author].length > 255) report.add("C23 comment author over 255 characters")

					const id = attributes.get(w("id")) ?? ""
					let last: string | undefined

					for (let child = doc.firstChild(node); child >= 0; child = doc.nextSibling(child)) {
						if (doc.kinds[child] === NodeKind.Element && doc.localName(child) === "p")
							last = attributesOf(part, child).get(`{${W14}}paraId`)
					}

					collected.commentIds.push(id)
					collected.commentParaIds.set(id, last)
					break
				}
			}

			if (REVISIONS.has(local)) {
				const author = attributes.get(w("author"))

				if (author !== undefined) collected.authors.add(author)
			}
		}

		if (uri === WP && local === "docPr") {
			report.check("C08", changed)

			const id = attributes.get("id") ?? ""
			const number = /^\d+$/.test(id) ? Number(id) : Number.NaN

			collected.docPrIds.push(id)

			if (!(number >= 1 && number <= 2147483646)) report.add("C08 docPr id out of range")
		}

		// A note or comment is a story of its own; so is a text box.
		const ownStory = uri === W && doc.parent(node) === root && ["footnote", "endnote", "comment"].includes(local)
		const textbox = uri === W && local === "txbxContent"
		const outer = story

		if (ownStory || textbox) story = { fields: [] }

		const cell = inCell || (uri === W && local === "tc")

		for (let child = doc.firstChild(node); child >= 0; child = doc.nextSibling(child)) {
			if (doc.kinds[child] === NodeKind.Element) visit(child, cell, inTextbox || textbox, false)
		}

		if (ownStory || textbox) {
			storyEnd(report, story)
			story = outer
		}
	}

	if (root >= 0) visit(root, false, false, false)
	if (storyPart) storyEnd(report, story)

	report.duplicates("C16 paraId duplicated in its part", paraIds)
	report.duplicates(
		"C09 revision id duplicated in its part",
		revisionIds.filter(id => id !== "")
	)
	report.add("C09 bookmark range unpaired", [...bookmarkIds.values()].filter(count => count !== 0).length)
	report.add("C09 move range unpaired", [...moveRanges.values()].filter(count => count !== 0).length)
	report.add("C09 permission range unpaired", [...permRanges.values()].filter(count => count !== 0).length)
	report.add("C09 move without a destination", [...moveNames.from].filter(name => !moveNames.to.has(name)).length)
	report.duplicates("C09 bookmark id duplicated in its part", bookmarkStarts)
}

// The rows of a table, through row-level content controls and custom XML.
function rowsOf(part: Part, table: number): number[] {
	const out: number[] = []
	const walk = (node: number): void => {
		for (const child of children(part.doc, node)) {
			if (part.ns.elementUri(child) !== W) continue

			const local = part.doc.localName(child)

			if (local === "tr") out.push(child)
			else if (local === "sdt" || local === "customXml" || local === "sdtContent") walk(child)
		}
	}

	walk(table)

	return out
}

// The cells of a row with their spans, and its gridBefore/gridAfter as numbers.
function cellsOf(part: Part, row: number): (number | { span: number })[] {
	const out: (number | { span: number })[] = []
	const { doc, ns } = part
	const walk = (node: number): void => {
		for (const child of children(doc, node)) {
			if (ns.elementUri(child) !== W) continue

			const local = doc.localName(child)

			if (local === "tc") {
				const properties = [...children(doc, child)].find(item => doc.localName(item) === "tcPr")
				const gridSpan =
					properties === undefined ? undefined : [...children(doc, properties)].find(item => doc.localName(item) === "gridSpan")
				const value = gridSpan === undefined ? 1 : Number(attributesOf(part, gridSpan).get(w("val")) ?? "1")

				out.push({ span: Number.isFinite(value) && value > 0 ? value : 1 })
			} else if (local === "trPr") {
				for (const item of children(doc, child)) {
					const name = doc.localName(item)

					if (name === "gridBefore" || name === "gridAfter") out.push(Number(attributesOf(part, item).get(w("val")) ?? "0") || 0)
				}
			} else if (local === "sdt" || local === "customXml" || local === "sdtContent") walk(child)
		}
	}

	walk(row)

	return out
}

// Whether the last block-level child of a cell (through content controls and custom XML) is a paragraph.
function endsInParagraph(part: Part, cell: number): boolean {
	let last = -1

	for (const child of children(part.doc, cell)) {
		if (part.ns.elementUri(child) === W && BLOCKS.has(part.doc.localName(child))) last = child
	}

	if (last < 0) return false

	const local = part.doc.localName(last)

	if (local === "p") return true
	if (local === "sdt") {
		const content = [...children(part.doc, last)].find(child => part.doc.localName(child) === "sdtContent")

		return content !== undefined && endsInParagraph(part, content)
	}

	return local === "customXml" && endsInParagraph(part, last)
}

// XML 1.0 Char: #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF].
function badCharacter(code: number): boolean {
	return (
		(code < 0x20 && code !== 0x9 && code !== 0xa && code !== 0xd) ||
		code === 0xfffe ||
		code === 0xffff ||
		(code >= 0xd800 && code <= 0xdfff)
	)
}

// Whether a revision element records a change of properties (w:rPr's w:ins, m:ctrlPr's w:ins/w:del around
// w:rPr: CT_MathCtrlIns) rather than wrapping runs.
function inProperties(doc: XmlDocument, node: number): boolean {
	for (let at = doc.parent(node); at >= 0; at = doc.parent(at)) {
		const local = doc.localName(at)

		if (local.endsWith("Pr")) return true
		if (local !== "ins" && local !== "del") return false
	}

	return false
}

function startsWith(bytes: Uint8Array, at: number, text: string): boolean {
	for (let k = 0; k < text.length; k++) if (bytes[at + k] !== text.charCodeAt(k)) return false

	return true
}

// Characters XML 1.0 does not allow (raw or as references), and carriage returns a parser turns into line feeds
// or spaces where they are content: in attribute values and in text that is not just markup whitespace (a
// prolog's CRLF is harmless).
function characters(bytes: Uint8Array, report: Report): void {
	let raw = 0
	let carriage = 0
	let inTag = false
	let quote = 0
	let pending = 0
	let content = false

	for (let at = 0; at < bytes.length; at++) {
		const byte = bytes[at] as number

		if (byte < 0x20 && byte !== 0x9 && byte !== 0xa && byte !== 0xd) raw++
		else if (byte === 0xef && bytes[at + 1] === 0xbf && (bytes[at + 2] === 0xbe || bytes[at + 2] === 0xbf)) raw++
		else if (byte === 0xed && (bytes[at + 1] ?? 0) >= 0xa0) raw++
		else if (byte === 0x26 && bytes[at + 1] === 0x23) {
			const end = bytes.indexOf(0x3b, at)
			const text = new TextDecoder().decode(bytes.subarray(at + 2, end < 0 ? at + 2 : end))
			const code = text.startsWith("x") ? Number.parseInt(text.slice(1), 16) : Number.parseInt(text, 10)

			if (badCharacter(code)) raw++
		}

		if (inTag) {
			if (quote !== 0) {
				if (byte === quote) quote = 0
				else if (byte === 0xd) carriage++
			} else if (byte === 0x22 || byte === 0x27) quote = byte
			else if (byte === 0x3e) inTag = false

			continue
		}

		if (byte === 0x3c) {
			if (content) carriage += pending

			pending = 0
			content = false

			// Comments, CDATA and processing instructions: skipped whole (CDATA is content).
			const close = startsWith(bytes, at, "<!--")
				? "-->"
				: startsWith(bytes, at, "<![CDATA[")
					? "]]>"
					: startsWith(bytes, at, "<?")
						? "?>"
						: ""

			if (close !== "") {
				const end = bytes.indexOf(close.charCodeAt(0), at + 2)
				let stop = end

				while (stop >= 0 && !startsWith(bytes, stop, close)) stop = bytes.indexOf(close.charCodeAt(0), stop + 1)

				const last = stop < 0 ? bytes.length : stop + close.length

				if (close === "]]>") for (let k = at; k < last; k++) if (bytes[k] === 0xd) carriage++

				at = last - 1
				continue
			}

			inTag = true
		} else if (byte === 0xd) pending++
		else if (byte !== 0x20 && byte !== 0x9 && byte !== 0xa) content = true
	}

	report.add("C20 character not allowed in XML 1.0", raw)
	report.add("C20 carriage return in content", carriage)
	report.check("C20", true)
}

function parsed(bytes: Uint8Array): XmlDocument | undefined {
	try {
		return parseXml(bytes)
	} catch {
		return undefined
	}
}

// Every violation of the catalogue in a package; `changed` limits what counts as checked (coverage).
export async function checkConformance(bytes: Uint8Array, changed?: ReadonlySet<string>): Promise<ConformanceReport> {
	const report = new Report()
	const pkg = await OpcPackage.open(bytesSource(bytes))
	const types = await pkg.contentTypes()
	const names = pkg.zip.entries.filter(entry => !entry.name.endsWith("/")).map(entry => entry.name)
	const existing = new Set(names.map(partKey))
	const isChanged = (name: string): boolean => changed === undefined || changed.has(name)
	const parts = new Map<string, Part>()
	const relationshipsOf = new Map<string, Map<string, Relationship>>()
	let mainPart: string | undefined

	// C21: content types.
	const contentTypes = parsed(await pkg.read("[Content_Types].xml")) ?? parseXml(new TextEncoder().encode("<Types/>"))
	const ctRoot = contentTypes.documentElement()
	const defaults: string[] = []
	const overrides: string[] = []

	for (let child = contentTypes.firstChild(ctRoot); child >= 0; child = contentTypes.nextSibling(child)) {
		if (contentTypes.kinds[child] !== NodeKind.Element) continue
		if (contentTypes.localName(child) === "Default") defaults.push((contentTypes.attribute(child, "Extension") ?? "").toLowerCase())
		if (contentTypes.localName(child) === "Override") overrides.push(partKey(contentTypes.attribute(child, "PartName") ?? ""))
	}

	report.duplicates("C21 Default extension duplicated", defaults)
	report.duplicates("C21 Override part duplicated", overrides)

	for (const name of names) {
		if (name === "[Content_Types].xml") continue

		report.check("C21", isChanged(name))

		if (types.typeOf(name) === undefined) report.add("C21 part without a content type")
	}

	// C22: relationships.
	for (const name of names) {
		if (types.typeOf(name) !== RELS_TYPE) continue

		report.check("C22", isChanged(name))

		const source = relsSource(name)
		const doc = parsed(await pkg.read(name))

		if (doc === undefined) continue

		const ns = new Namespaces(doc)
		const ids: string[] = []
		const map = new Map<string, Relationship>()

		if (source !== "" && !existing.has(partKey(source))) report.add("C22 relationships of a missing part")

		for (const node of children(doc, doc.documentElement())) {
			if (ns.elementUri(node) !== PACKAGE_RELS || doc.localName(node) !== "Relationship") continue

			const id = doc.attribute(node, "Id") ?? ""
			const type = doc.attribute(node, "Type") ?? ""
			const target = doc.attribute(node, "Target") ?? ""
			const mode = doc.attribute(node, "TargetMode")
			const external = mode === "External"

			ids.push(id)
			map.set(id, { type, external, target: external ? target : resolveTarget(source, target) })

			if (mode !== undefined && mode !== "External" && mode !== "Internal") report.add("C22 TargetMode neither External nor Internal")
			if (type === HYPERLINK && !external && /^[a-z][a-z0-9+.-]*:/i.test(target)) report.add("C13 hyperlink to a URL not External")
			if (!external && !/^[a-z][a-z0-9+.-]*:/i.test(target)) {
				const resolved = resolveTarget(source, target.replace(/#.*$/, ""))

				if (resolved !== "" && !existing.has(partKey(resolved))) report.add("C22 internal target missing")
			}

			if (source === "" && type.endsWith("/officeDocument")) mainPart = resolveTarget("", target)
		}

		report.duplicates("C22 relationship id duplicated", ids)
		relationshipsOf.set(partKey(source), map)
	}

	// Parse every XML part with a WordprocessingML or properties type, and check characters of the changed ones.
	for (const name of names) {
		const type = types.typeOf(name) ?? ""
		const xmlPart = type.endsWith("+xml") || type === RELS_TYPE || name === "[Content_Types].xml"

		if (!xmlPart) continue

		const data = await pkg.read(name)

		if (isChanged(name)) characters(data, report)

		const doc = parsed(data)

		if (doc === undefined) {
			report.add("C20 part not well-formed XML")
			continue
		}

		if (PROPERTIES.has(type)) {
			report.check("C19", isChanged(name))

			const ns = new Namespaces(doc)

			for (let node = 0; node < doc.count; node++) {
				if (doc.kinds[node] !== NodeKind.Element) continue
				if (ns.elementUri(node) === MC) report.add("C19 markup compatibility in a properties part")

				for (const attribute of doc.attributes(node)) {
					const colon = attribute.name.indexOf(":")

					if (colon > 0 && !attribute.name.startsWith("xmlns:") && ns.uri(node, attribute.name.slice(0, colon)) === MC)
						report.add("C19 markup compatibility in a properties part")
				}
			}
		}

		if (!type.startsWith(WML) && !MACRO_MAIN.test(type)) continue

		parts.set(partKey(name), { name, type, doc, ns: new Namespaces(doc), changed: isChanged(name) })
	}

	// The glossary document (building blocks) is a document of its own: ids are unique within each.
	const glossary = [...parts.values()].find(part => part.type === `${WML}document.glossary+xml`)
	const glossaryParts = new Set<string>()

	if (glossary !== undefined) {
		glossaryParts.add(partKey(glossary.name))

		for (const relationship of relationshipsOf.get(partKey(glossary.name))?.values() ?? []) {
			if (!relationship.external) glossaryParts.add(partKey(relationship.target))
		}
	}

	for (const inGlossary of glossary === undefined ? [false] : [false, true]) {
		const scope = new Map([...parts].filter(([key]) => glossaryParts.has(key) === inGlossary))
		const main = inGlossary ? glossary?.name : mainPart

		documentChecks(scope, report, relationshipsOf, main)
	}

	return report
}

// The checks across the parts of one document (the main one or the glossary).
function documentChecks(
	parts: Map<string, Part>,
	report: Report,
	relationshipsOf: Map<string, Map<string, Relationship>>,
	mainPart: string | undefined
): void {
	const collected: Collected = {
		docPrIds: [],
		bookmarkNames: [],
		sdtIds: [],
		commentRefs: [],
		commentRanges: new Map(),
		footnoteRefs: [],
		endnoteRefs: [],
		numIdRefs: [],
		authors: new Set(),
		commentParaIds: new Map(),
		commentIds: [],
		commentRangeCount: 0
	}

	for (const part of parts.values()) walkPart(part, report, collected, relationshipsOf.get(partKey(part.name)) ?? new Map())

	// Package-wide uniqueness.
	report.duplicates("C08 docPr id duplicated", collected.docPrIds)
	report.duplicates(
		"C14 bookmark name duplicated",
		collected.bookmarkNames.filter(name => name !== "_GoBack")
	)
	report.duplicates("C15 sdt id duplicated", collected.sdtIds)

	// Comments.
	const commentIds = new Set(collected.commentIds)

	report.duplicates("C09 comment id duplicated", collected.commentIds)
	report.add("C09 comment reference unresolved", collected.commentRefs.filter(id => !commentIds.has(id)).length)
	report.add("C09 comment range unpaired", [...collected.commentRanges.values()].filter(count => count !== 0).length)

	if (collected.commentRangeCount > 32767) report.add("C09 more than 32767 comment ranges")

	// C06: notes.
	for (const [kind, refs] of [
		["footnote", collected.footnoteRefs],
		["endnote", collected.endnoteRefs]
	] as const) {
		const part = [...parts.values()].find(item => item.type === `${WML}${kind}s+xml`)

		if (part === undefined) {
			report.add(`C06 ${kind} reference without a notes part`, refs.length)
			continue
		}

		report.check("C06", part.changed)

		const ids: string[] = []
		const normal = new Set<string>()
		const separators = new Set<string>()
		const root = part.doc.documentElement()

		for (const note of children(part.doc, root)) {
			if (part.doc.localName(note) !== kind) continue

			const attributes = attributesOf(part, note)
			const id = attributes.get(w("id")) ?? ""
			const type = attributes.get(w("type")) ?? "normal"

			ids.push(id)

			if (!/^-?\d+$/.test(id) || Number(id) < -2147483648 || Number(id) > 32767) report.add(`C06 ${kind} id out of range`)
			if (type === "normal") normal.add(id)
			else separators.add(type)
		}

		report.duplicates(`C06 ${kind} id duplicated`, ids)

		if (!separators.has("separator") || !separators.has("continuationSeparator")) report.add(`C06 ${kind}s part without separators`)

		report.add(`C06 ${kind} reference unresolved`, refs.filter(id => !normal.has(id)).length)
	}

	// C07: numbering.
	const numbering = [...parts.values()].find(item => item.type === `${WML}numbering+xml`)

	if (numbering !== undefined) {
		report.check("C07", numbering.changed)

		const abstractIds: string[] = []
		const numIds: string[] = []
		const nsids: string[] = []
		const references: string[] = []
		let numSeen = false

		for (const node of children(numbering.doc, numbering.doc.documentElement())) {
			if (numbering.ns.elementUri(node) !== W) continue

			const local = numbering.doc.localName(node)
			const attributes = attributesOf(numbering, node)

			if (local === "abstractNum") {
				const id = attributes.get(w("abstractNumId")) ?? ""

				if (numSeen) report.add("C07 abstractNum after num")
				if (!/^\d+$/.test(id)) report.add("C07 abstractNumId negative or not a number")

				abstractIds.push(id)

				for (const child of children(numbering.doc, node)) {
					if (numbering.doc.localName(child) === "nsid") nsids.push(attributesOf(numbering, child).get(w("val")) ?? "")
				}
			} else if (local === "num") {
				numSeen = true
				numIds.push(attributes.get(w("numId")) ?? "")

				for (const child of children(numbering.doc, node)) {
					if (numbering.doc.localName(child) === "abstractNumId")
						references.push(attributesOf(numbering, child).get(w("val")) ?? "")
				}
			}
		}

		const abstract = new Set(abstractIds)
		const nums = new Set(numIds)

		report.duplicates("C07 abstractNumId duplicated", abstractIds)
		report.duplicates("C07 numId duplicated", numIds)
		report.duplicates("C07 nsid duplicated", nsids)
		report.add("C07 num refers to a missing abstractNum", references.filter(id => !abstract.has(id)).length)
		report.add("C07 numId unresolved", collected.numIdRefs.filter(id => id !== "0" && !nums.has(id)).length)
	} else report.add("C07 numId unresolved", collected.numIdRefs.filter(id => id !== "0").length)

	// C17: Word's comment extension parts.
	const main = mainPart === undefined ? undefined : relationshipsOf.get(partKey(mainPart))
	const extension = (local: string): Part | undefined => [...parts.values()].find(item => item.type === `${WML}${local}+xml`)
	const commentParaIds = new Set([...collected.commentParaIds.values()].filter((id): id is string => id !== undefined))

	for (const local of Object.keys(EXTENSION_RELATIONSHIPS)) {
		const part = extension(local)

		if (part === undefined) continue

		report.check("C17", part.changed)

		if (![...(main?.values() ?? [])].some(relationship => relationship.type === EXTENSION_RELATIONSHIPS[local]))
			report.add(`C17 ${local} part not related`)
	}

	const extended = extension("commentsExtended")
	const exParaIds = new Set<string>()

	if (extended !== undefined) {
		const parents: string[] = []

		for (const node of children(extended.doc, extended.doc.documentElement())) {
			const attributes = attributesOf(extended, node)
			const paraId = w15(attributes, "paraId") ?? ""
			const parent = w15(attributes, "paraIdParent")

			exParaIds.add(paraId)

			if (!commentParaIds.has(paraId)) report.add("C17 commentEx paraId not a comment's last paragraph")
			if (parent !== undefined) parents.push(parent)
		}

		report.add("C17 commentEx parent unresolved", parents.filter(parent => !exParaIds.has(parent)).length)
	}

	const ids = extension("commentsIds")
	const durableIds = new Set<string>()

	if (ids !== undefined) {
		const all: string[] = []

		for (const node of children(ids.doc, ids.doc.documentElement())) {
			const attributes = attributesOf(ids, node)
			const durable = attributes.get(`{${W16CID}}durableId`) ?? ""
			const number = /^[0-9A-Fa-f]{1,8}$/.test(durable) ? Number.parseInt(durable, 16) : Number.NaN

			all.push(durable)
			durableIds.add(durable)

			if (!(number > 0 && number < 0x7fffffff)) report.add("C17 durableId out of range")
			if (!commentParaIds.has(attributes.get(`{${W16CID}}paraId`) ?? ""))
				report.add("C17 commentId paraId not a comment's last paragraph")
		}

		report.duplicates("C17 durableId duplicated", all)
	}

	const extensible = extension("commentsExtensible")

	if (extensible !== undefined) {
		for (const node of children(extensible.doc, extensible.doc.documentElement())) {
			const attributes = attributesOf(extensible, node)

			if (extensible.doc.localName(node) !== "commentExtensible") continue
			if (!durableIds.has(attributes.get(`{${W16CEX}}durableId`) ?? "")) report.add("C17 commentExtensible durableId unresolved")

			const date = attributes.get(`{${W16CEX}}dateUtc`)

			if (date !== undefined && !WORD_DATE.test(date)) report.add("C01 dateUtc not UTC")
		}
	}

	const people = extension("people")

	if (people !== undefined) {
		for (const node of children(people.doc, people.doc.documentElement())) {
			const author = w15(attributesOf(people, node), "author")

			if (author !== undefined && !collected.authors.has(author)) report.add("C17 person without a comment or revision")
		}
	}
}

// Violations that are a value of copied markup, not a relation between constructs.
const CARRIED: readonly string[] = [
	"C01 annotation date",
	"C02 ",
	"C03 ",
	"C04 ",
	"C05 ",
	"C13 hyperlink anchor",
	"C13 hyperlink docLocation",
	"C13 hyperlink tgtFrame",
	"C13 hyperlink tooltip",
	"C20 character",
	"C20 carriage return"
]

// Violations of `after` beyond `before`'s counts, as "key (before -> after)".
export function newViolations(before: ConformanceReport, after: ConformanceReport): string[] {
	const out: string[] = []

	for (const [key, count] of after.violations) {
		const was = before.violations.get(key) ?? 0

		// A value the source already had wrong is repeated by content copied with it (a split run, a table split
		// or a pasted row keeps its properties): only a first one is the writer's.
		if (was > 0 && CARRIED.some(prefix => key.startsWith(prefix))) continue
		if (count > was) out.push(`${key} (${was} -> ${count})`)
	}

	return out
}
