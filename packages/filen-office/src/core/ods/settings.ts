import type { OdfPackage } from "../odf/package"
import { canonicalAttributes, elementName } from "../odf/tree"
import type { SliceClock } from "../slicing"
import { Namespaces, NodeKind, XmlError, type XmlDocument } from "../xml/index"

// settings.xml (ODF 1.3 Part 3 §3.10): LibreOffice's view settings, config:config-item-set "ooo:view-settings" →
// "Views" → the first view's items and its per-sheet "Tables" map. Only what the sheet view shows is read; a
// package without the stream, or a stream that does not read, has none.

export interface OdsTableView {
	readonly cursorCol: number
	readonly cursorRow: number
	// Split modes: 0 none, 1 split, 2 frozen (LibreOffice's ScSplitMode).
	readonly horizontalSplitMode: number
	readonly verticalSplitMode: number
	// For a frozen split, the first column and row past the frozen ones.
	readonly horizontalSplitPosition: number
	readonly verticalSplitPosition: number
	readonly positionLeft: number
	readonly positionRight: number
	readonly positionTop: number
	readonly positionBottom: number
	readonly zoom: number
	readonly showGrid: boolean | undefined
}

export interface OdsSettings {
	readonly activeTable: string | undefined
	readonly tables: ReadonlyMap<string, OdsTableView>
	readonly showGrid: boolean
	readonly headings: boolean
}

const NONE: OdsSettings = { activeTable: undefined, tables: new Map(), showGrid: true, headings: true }

function children(document: XmlDocument, node: number): number[] {
	const out: number[] = []

	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) out.push(child)

	return out
}

export async function readSettings(odf: OdfPackage, clock: SliceClock | undefined): Promise<OdsSettings> {
	let document: XmlDocument | undefined

	try {
		document = await odf.xml("settings.xml", clock)
	} catch (error) {
		if (error instanceof XmlError) return NONE

		throw error
	}

	if (document === undefined) return NONE

	const namespaces = new Namespaces(document, { lenient: true })
	const name = (node: number): string => elementName(document, namespaces, node)
	const configName = (node: number): string | undefined => canonicalAttributes(document, namespaces, node).get("config:name")
	// The items of a set or entry by name (config:config-item), as text.
	const items = (node: number): Map<string, string> => {
		const out = new Map<string, string>()

		for (const child of children(document, node)) {
			if (name(child) !== "config:config-item") continue

			const key = configName(child)

			if (key !== undefined)
				out.set(key, document.kind(document.firstChild(child)) === NodeKind.Text ? document.textContent(child) : "")
		}

		return out
	}
	const root = document.documentElement()
	const settings = children(document, root).find(node => name(node) === "office:settings")

	if (settings === undefined) return NONE

	const viewSet = children(document, settings).find(
		node => name(node) === "config:config-item-set" && configName(node) === "ooo:view-settings"
	)
	const views = viewSet === undefined ? undefined : children(document, viewSet).find(node => configName(node) === "Views")
	const view = views === undefined ? undefined : children(document, views).find(node => name(node) === "config:config-item-map-entry")

	if (view === undefined) return NONE

	const viewItems = items(view)
	const tables = new Map<string, OdsTableView>()
	const tablesMap = children(document, view).find(node => name(node) === "config:config-item-map-named" && configName(node) === "Tables")
	const number = (map: ReadonlyMap<string, string>, key: string, fallback: number): number => {
		const value = Number(map.get(key))

		return Number.isFinite(value) ? value : fallback
	}

	for (const entry of tablesMap === undefined ? [] : children(document, tablesMap)) {
		const table = configName(entry)

		if (table === undefined) continue

		const values = items(entry)
		const grid = values.get("ShowGrid")

		tables.set(table, {
			cursorCol: number(values, "CursorPositionX", 0),
			cursorRow: number(values, "CursorPositionY", 0),
			horizontalSplitMode: number(values, "HorizontalSplitMode", 0),
			verticalSplitMode: number(values, "VerticalSplitMode", 0),
			horizontalSplitPosition: number(values, "HorizontalSplitPosition", 0),
			verticalSplitPosition: number(values, "VerticalSplitPosition", 0),
			positionLeft: number(values, "PositionLeft", 0),
			positionRight: number(values, "PositionRight", 0),
			positionTop: number(values, "PositionTop", 0),
			positionBottom: number(values, "PositionBottom", 0),
			zoom: number(values, "ZoomValue", 100),
			showGrid: grid === undefined ? undefined : grid === "true"
		})
	}

	return {
		activeTable: viewItems.get("ActiveTable"),
		tables,
		showGrid: viewItems.get("ShowGrid") !== "false",
		headings: viewItems.get("HasColumnRowHeaders") !== "false"
	}
}
