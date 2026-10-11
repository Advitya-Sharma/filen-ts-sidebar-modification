import { OpcError, ZipError } from "../../opc/index"
import { XmlError, type XmlDocument } from "../../xml/index"
import type { DocxModel } from "../model/index"
import { parseNumbering, type Numbering } from "../numbering/index"
import { gateReasons, type DocxGateReason, type DocxPackage } from "../package"
import { PropsStore } from "../props/index"
import { DocxStyles, PropertyResolver, parseStyles, styleSettings } from "../styles/index"
import { parseTheme } from "../theme/index"
import type { NewList } from "./formats"
import { RunBases } from "./formatting"

// What editing one document needs besides the model: the shared property store (so run bases intern
// with everything else), styles, numbering and the resolver for the formatting state of a selection.
export interface DocxEditContext {
	readonly model: DocxModel
	readonly store: PropsStore
	readonly styles: DocxStyles
	readonly numbering: Numbering | undefined
	readonly resolver: PropertyResolver
	readonly bases: RunBases
	// Gate 0: why the document opens view-only (empty when it may be edited).
	readonly gate: readonly DocxGateReason[]
	// Lists commands started, by their (negative) id, shared by every session of the document.
	readonly newLists: Map<number, NewList>
	// Built-in paragraph styles commands used that the document lacks: id -> built-in name.
	readonly newStyles: Map<string, string>
}

function isDocumentError(error: unknown): boolean {
	return error instanceof XmlError || error instanceof ZipError || error instanceof OpcError
}

// A part read for editing that exists but cannot be parsed makes the document view-only.
async function optionalXml(pkg: DocxPackage, kind: string, unreadable: string[]): Promise<XmlDocument | undefined> {
	const name = pkg.part(kind)

	if (name === undefined || !pkg.has(name)) return undefined

	try {
		return await pkg.xml(name)
	} catch (error) {
		if (!isDocumentError(error)) throw error

		unreadable.push(name)

		return undefined
	}
}

export interface EditContextParts {
	readonly store?: PropsStore
	readonly styles?: DocxStyles
	readonly numbering?: Numbering
	readonly resolver?: PropertyResolver
}

// Reads styles, numbering and theme for `model` (or reuses what a caller already parsed for layout).
export async function openDocxEditContext(pkg: DocxPackage, model: DocxModel, reuse: EditContextParts = {}): Promise<DocxEditContext> {
	const unreadable: string[] = []
	const store = reuse.store ?? new PropsStore()
	let styles = reuse.styles
	let numbering = reuse.numbering
	let resolver = reuse.resolver

	if (resolver === undefined) {
		const stylesXml = styles === undefined ? await optionalXml(pkg, "styles", unreadable) : undefined

		styles ??= stylesXml === undefined ? DocxStyles.empty() : parseStyles(stylesXml, store)

		if (numbering === undefined) {
			const numberingXml = await optionalXml(pkg, "numbering", unreadable)

			numbering = numberingXml === undefined ? undefined : parseNumbering(numberingXml, store, styles)
		}

		const themeXml = await optionalXml(pkg, "theme", unreadable)
		const theme = themeXml === undefined ? undefined : parseTheme(themeXml)

		resolver = new PropertyResolver({ styles, numbering, theme, settings: styleSettings(model.settings) })
	}

	const gate = gateReasons(pkg.facts, model, resolver.styles, resolver.numbering, {
		malformed: unreadable.length > 0 || store.malformed.length > 0
	})

	return {
		model,
		store,
		styles: resolver.styles,
		numbering: resolver.numbering,
		resolver,
		bases: new RunBases(store),
		gate,
		newLists: new Map(),
		newStyles: new Map()
	}
}
