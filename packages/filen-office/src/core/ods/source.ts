import type { OdfPackage } from "../odf/package"
import type { ZipEntry } from "../opc/zipReader"
import type { Relationship } from "../opc/relationships"
import type { SheetId } from "../sheet/protocol"
import type {
	XlsxColRange,
	XlsxDefinedName,
	XlsxSection,
	XlsxSheetProtection,
	XlsxSheetView,
	XlsxWorkbookModel,
	XlsxWorkbookProtection
} from "../xlsx/model/types"
import type { OdsIssue } from "./issues"
import type { OdsUnheld } from "./unheld"
import type { OdsSettings } from "./settings"
import type { OdsCellStyles } from "./styles/cell"
import type { OdsChartSource } from "./charts/read"

// What a save needs of the package an OpenDocument workbook was read from, kept beside the model: the archive to
// copy from and stream content.xml out of again, the styles the cells' formats came from, and each table as the
// reader built it, so the writer tells what an edit changed.

export interface OdsTableSource {
	// Its position among content.xml's tables, and the sheet it became.
	readonly index: number
	readonly id: SheetId
	readonly name: string
	readonly styleName: string | undefined
	readonly hidden: boolean
	readonly rightToLeft: boolean
	readonly tabColor: number
	readonly cols: readonly XlsxColRange[]
	readonly merges: Uint32Array
	readonly sections: readonly XlsxSection[]
	readonly relationships: readonly Relationship[]
	readonly view: XlsxSheetView
	readonly protection: XlsxSheetProtection | null
	readonly defaultRowHeight: number
	readonly defaultColPx: number
	// Content past what is materialised (a repeat beyond the caps): its rows are never rewritten.
	readonly truncated: boolean
	// The rows its row elements give and the columns its column elements declare.
	readonly rows: number
	readonly columns: number
	// Its validations' rules (dataValidation without sqref) to the table:content-validation names they came from.
	readonly validationNames: ReadonlyMap<string, string>
}

export interface OdsSource {
	readonly odf: OdfPackage
	// office:version of content.xml's root, which the package's other parts are of too.
	readonly version: string | undefined
	// The table:content-validation names content.xml declares, which new ones must not take.
	readonly validationNames: ReadonlySet<string>
	readonly content: ZipEntry
	// The root's namespace declarations (prefix to URI, "" for a default namespace).
	readonly prefixes: ReadonlyMap<string, string>
	readonly styles: OdsCellStyles
	readonly settings: OdsSettings
	readonly tables: readonly OdsTableSource[]
	// The features' parts as built from content.xml (notes, charts), by name.
	readonly parts: ReadonlyMap<string, Uint8Array>
	// Which ODF object each chart part came from.
	readonly charts: readonly OdsChartSource[]
	readonly definedNames: readonly XlsxDefinedName[]
	readonly protection: XlsxWorkbookProtection | null
	readonly activeTab: number
	readonly issues: readonly OdsIssue[]
	// What names cells by address that the model does not carry.
	readonly unheld: OdsUnheld
}

const sources = new WeakMap<XlsxWorkbookModel, OdsSource>()

export function setOdsSource(workbook: XlsxWorkbookModel, source: OdsSource): void {
	sources.set(workbook, source)
}

export function odsSource(workbook: XlsxWorkbookModel): OdsSource | undefined {
	return sources.get(workbook)
}
