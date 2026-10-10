import type { XmlDocument } from "../../xml/index"
import { EMPTY_PARAGRAPH, EMPTY_RUN } from "../props/parse"
import { PropsReader } from "../props/reader"
import type { PropsStore } from "../props/store"
import type { ParagraphProps, RunProps } from "../props/types"
import type { DocxStyles } from "../styles/styles"

// word/numbering.xml (ECMA-376 Part 1 §17.9).

export const MAX_LEVEL = 8

export interface NumberingLevel {
	readonly ilvl: number
	// w:start; zero when absent (§17.9.25).
	readonly start: number
	// ST_NumberFormat ("decimal", "bullet", "lowerRoman"...).
	readonly format: string
	// w:numFmt/@w:format for "custom".
	readonly customFormat: string | undefined
	// w:lvlRestart: undefined = restart after any higher level, 0 = never, n = after level n (1-based).
	readonly restart: number | undefined
	readonly styleId: string | undefined
	readonly legal: boolean
	readonly suffix: "tab" | "space" | "nothing"
	// w:lvlText; undefined when absent or w:null.
	readonly text: string | undefined
	readonly pictureBulletId: number | undefined
	// ST_Jc: "start"/"left", "center", "end"/"right".
	readonly alignment: string
	readonly tentative: boolean
	readonly paragraph: ParagraphProps
	readonly run: RunProps
	// The w:lvl element in numbering.xml.
	readonly node: number
}

export interface AbstractNumbering {
	readonly id: number
	readonly nsid: string | undefined
	readonly multiLevelType: string | undefined
	readonly template: string | undefined
	readonly name: string | undefined
	readonly styleLink: string | undefined
	readonly numStyleLink: string | undefined
	readonly levels: readonly (NumberingLevel | undefined)[]
	readonly node: number
}

export interface LevelOverride {
	readonly start: number | undefined
	readonly level: NumberingLevel | undefined
}

export interface NumberingInstance {
	readonly numId: number
	readonly abstractId: number
	readonly overrides: ReadonlyMap<number, LevelOverride>
	readonly node: number
}

// w:numPicBullet: the picture lives in a w:pict (VML) or w:drawing child of `node`.
export interface PictureBullet {
	readonly id: number
	readonly node: number
	// r:id / r:embed of the image, when it could be found.
	readonly relationshipId: string | undefined
}

// The definition a (numId, ilvl) pair uses, after numStyleLink indirection and overrides.
export interface EffectiveLevel {
	readonly numId: number
	readonly ilvl: number
	readonly level: NumberingLevel
	// Counters are shared by every w:num of one abstract definition.
	readonly abstractId: number
	readonly start: number
	readonly startOverridden: boolean
}

function clampLevel(ilvl: number): number {
	return ilvl < 0 ? 0 : ilvl > MAX_LEVEL ? MAX_LEVEL : ilvl
}

export class Numbering {
	readonly document: XmlDocument | undefined
	readonly abstracts: ReadonlyMap<number, AbstractNumbering>
	readonly instances: ReadonlyMap<number, NumberingInstance>
	readonly pictureBullets: ReadonlyMap<number, PictureBullet>
	// A start tag in numbering.xml could not be read or repeats an attribute: the document is view-only.
	readonly malformed: boolean
	private readonly styles: DocxStyles | undefined
	private readonly cache = new Map<number, EffectiveLevel | null>()

	constructor(
		document: XmlDocument | undefined,
		abstracts: ReadonlyMap<number, AbstractNumbering>,
		instances: ReadonlyMap<number, NumberingInstance>,
		pictureBullets: ReadonlyMap<number, PictureBullet>,
		styles: DocxStyles | undefined,
		malformed: boolean = false
	) {
		this.malformed = malformed
		this.document = document
		this.abstracts = abstracts
		this.instances = instances
		this.pictureBullets = pictureBullets
		this.styles = styles
	}

	static empty(): Numbering {
		return new Numbering(undefined, new Map(), new Map(), new Map(), undefined)
	}

	has(numId: number): boolean {
		return numId > 0 && this.instances.has(numId)
	}

	// The abstract definition holding the levels: an abstractNum with w:numStyleLink takes its levels from
	// the numbering style it names (via that style's numId), at most a few hops and never in a cycle.
	abstractFor(instance: NumberingInstance): AbstractNumbering | undefined {
		let abstract = this.abstracts.get(instance.abstractId)
		const seen = new Set<number>()

		while (abstract?.numStyleLink !== undefined && this.styles !== undefined && !seen.has(abstract.id)) {
			seen.add(abstract.id)

			const style = this.styles.byId(abstract.numStyleLink)
			const linkedNumId = style?.type === "numbering" ? style.paragraph.numbering?.numId : undefined
			const linked = linkedNumId === undefined ? undefined : this.instances.get(linkedNumId)
			const target = linked === undefined ? undefined : this.abstracts.get(linked.abstractId)

			if (target === undefined) break

			abstract = target
		}

		return abstract
	}

	level(numId: number, ilvl: number): EffectiveLevel | undefined {
		const lvl = clampLevel(ilvl)
		const key = numId * 16 + lvl
		const cached = this.cache.get(key)

		if (cached !== undefined) return cached ?? undefined

		const instance = numId > 0 ? this.instances.get(numId) : undefined
		let result: EffectiveLevel | null = null

		if (instance !== undefined) {
			const abstract = this.abstractFor(instance)
			const override = instance.overrides.get(lvl)
			const level = override?.level ?? abstract?.levels[lvl]

			if (level !== undefined && abstract !== undefined) {
				result = {
					numId,
					ilvl: lvl,
					level,
					abstractId: abstract.id,
					start: override?.start ?? level.start,
					startOverridden: override?.start !== undefined
				}
			}
		}

		this.cache.set(key, result)

		return result ?? undefined
	}
}

function readLevel(r: PropsReader, store: PropsStore, node: number): NumberingLevel | undefined {
	const document = r.document
	const ilvl = r.integer(node, "ilvl")

	if (ilvl === undefined || ilvl < 0 || ilvl > MAX_LEVEL) return undefined

	let start = 0
	let format = "decimal"
	let customFormat: string | undefined
	let restart: number | undefined
	let styleId: string | undefined
	let legal = false
	let suffix: "tab" | "space" | "nothing" = "tab"
	let text: string | undefined
	let pictureBulletId: number | undefined
	let alignment = "start"
	let paragraph = EMPTY_PARAGRAPH
	let run = EMPTY_RUN

	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
		switch (r.local(child)) {
			case "start":
				start = r.integer(child, "val") ?? 0
				break
			case "numFmt":
				format = r.string(child, "val") ?? "decimal"
				customFormat = r.string(child, "format")
				break
			case "lvlRestart":
				restart = r.integer(child, "val")
				break
			case "pStyle":
				styleId = r.string(child, "val")
				break
			case "isLgl":
				legal = r.toggle(child)
				break
			case "suff": {
				const value = r.string(child, "val")

				suffix = value === "space" || value === "nothing" ? value : "tab"

				break
			}
			case "lvlText":
				text = r.onOff(child, "null", false) ? undefined : r.string(child, "val")
				break
			case "lvlPicBulletId":
				pictureBulletId = r.integer(child, "val")
				break
			case "lvlJc":
				alignment = r.string(child, "val") ?? "start"
				break
			case "pPr":
				paragraph = store.paragraph(document, child)
				break
			case "rPr":
				run = store.run(document, child)
				break
		}
	}

	return {
		ilvl,
		start,
		format,
		customFormat,
		restart,
		styleId,
		legal,
		suffix,
		text,
		pictureBulletId,
		alignment,
		tentative: r.onOff(node, "tentative", false),
		paragraph,
		run,
		node
	}
}

function readAbstract(r: PropsReader, store: PropsStore, node: number, id: number): AbstractNumbering {
	const document = r.document
	const levels: (NumberingLevel | undefined)[] = new Array<NumberingLevel | undefined>(MAX_LEVEL + 1).fill(undefined)
	let nsid: string | undefined
	let multiLevelType: string | undefined
	let template: string | undefined
	let name: string | undefined
	let styleLink: string | undefined
	let numStyleLink: string | undefined

	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
		switch (r.local(child)) {
			case "nsid":
				nsid = r.string(child, "val")
				break
			case "multiLevelType":
				multiLevelType = r.string(child, "val")
				break
			case "tmpl":
				template = r.string(child, "val")
				break
			case "name":
				name = r.string(child, "val")
				break
			case "styleLink":
				styleLink = r.string(child, "val")
				break
			case "numStyleLink":
				numStyleLink = r.string(child, "val")
				break
			case "lvl": {
				const level = readLevel(r, store, child)

				if (level !== undefined && levels[level.ilvl] === undefined) levels[level.ilvl] = level

				break
			}
		}
	}

	return { id, nsid, multiLevelType, template, name, styleLink, numStyleLink, levels, node }
}

function readInstance(r: PropsReader, store: PropsStore, node: number, numId: number): NumberingInstance | undefined {
	const document = r.document
	const abstractNode = r.child(node, "abstractNumId")
	const abstractId = abstractNode < 0 ? undefined : r.integer(abstractNode, "val")

	if (abstractId === undefined) return undefined

	const overrides = new Map<number, LevelOverride>()

	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
		if (r.local(child) !== "lvlOverride") continue

		const ilvl = r.integer(child, "ilvl")

		if (ilvl === undefined || ilvl < 0 || ilvl > MAX_LEVEL) continue

		const startNode = r.child(child, "startOverride")
		const levelNode = r.child(child, "lvl")

		overrides.set(ilvl, {
			start: startNode < 0 ? undefined : r.integer(startNode, "val"),
			level: levelNode < 0 ? undefined : readLevel(r, store, levelNode)
		})
	}

	return { numId, abstractId, overrides, node }
}

// The image relationship of a picture bullet: a:blip/@r:embed (DrawingML) or v:imagedata/@r:id (VML).
function pictureRelationship(r: PropsReader, node: number): string | undefined {
	const document = r.document
	const end = document.end(node)

	for (let n = node + 1; n < document.count && document.start(n) < end; n++) {
		const local = document.localName(n)

		if (local === "blip") {
			const id = r.relationshipId(n, "embed")

			if (id !== undefined) return id
		} else if (local === "imagedata") {
			const id = r.relationshipId(n, "id")

			if (id !== undefined) return id
		}
	}

	return undefined
}

export function parseNumbering(document: XmlDocument, store: PropsStore, styles?: DocxStyles): Numbering {
	const r = PropsReader.for(document)
	// Checked first, so the property reads that follow find every start tag already validated.
	const malformed = r.model.sweepAll()
	const root = document.documentElement()

	if (root < 0 || r.local(root) !== "numbering") return new Numbering(undefined, new Map(), new Map(), new Map(), undefined, malformed)

	const abstracts = new Map<number, AbstractNumbering>()
	const instances = new Map<number, NumberingInstance>()
	const pictureBullets = new Map<number, PictureBullet>()

	for (let child = document.firstChildElement(root); child >= 0; child = document.nextSiblingElement(child)) {
		switch (r.local(child)) {
			case "numPicBullet": {
				const id = r.integer(child, "numPicBulletId")

				if (id !== undefined) pictureBullets.set(id, { id, node: child, relationshipId: pictureRelationship(r, child) })

				break
			}
			case "abstractNum": {
				const id = r.integer(child, "abstractNumId")

				if (id !== undefined && !abstracts.has(id)) abstracts.set(id, readAbstract(r, store, child, id))

				break
			}
			case "num": {
				const numId = r.integer(child, "numId")
				const instance = numId === undefined ? undefined : readInstance(r, store, child, numId)

				if (instance !== undefined && !instances.has(instance.numId)) instances.set(instance.numId, instance)

				break
			}
		}
	}

	return new Numbering(document, abstracts, instances, pictureBullets, styles, malformed)
}
