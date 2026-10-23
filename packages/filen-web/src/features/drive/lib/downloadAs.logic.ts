import { FileTextIcon, type LucideIcon } from "lucide-react"
import { asDirectoryOrFile, type DriveItem } from "@/features/drive/lib/item"
import { canPreview, previewType } from "@/features/drive/lib/preview.logic"
import { type DriveKey } from "@/lib/i18n"

// Download as in the item menus, the parts without a DOM: which files can be exported and to what. A file is
// exported from the saved file, laid out headless as its viewer lays it out, and downloaded; the drive is unchanged.

// The engine a file is exported from: the document viewer's (Word, Word 97–2003, OpenDocument text, RTF) or the
// spreadsheet viewer's (Excel, Excel 97–2003, OpenDocument spreadsheet, CSV and TSV).
export type DownloadAsSource = "document" | "spreadsheet"

export type DownloadAsTarget = "pdf"

export interface DownloadAsTargetDef {
	readonly labelKey: DriveKey
	readonly icon: LucideIcon
	// The extension the download is named with.
	readonly extension: string
	readonly sources: ReadonlySet<DownloadAsSource>
}

export const DOWNLOAD_AS_TARGETS: Readonly<Record<DownloadAsTarget, DownloadAsTargetDef>> = {
	pdf: { labelKey: "driveDownloadAsPdf", icon: FileTextIcon, extension: "pdf", sources: new Set(["document", "spreadsheet"]) }
}

// The submenu's entries in groups, a rule between groups (the PDF, then the editable formats a file converts to).
const DOWNLOAD_AS_GROUPS: readonly (readonly DownloadAsTarget[])[] = [["pdf"]]

// What a file is exported from; null for anything the viewers do not open: a directory, an undecryptable file, a
// file of another kind, or one past its viewer's size cap (bufferedSizeCap: the export holds the same whole buffer
// and lays it out as the viewer would, so a file the viewer refuses to load is never loaded for an export either).
export function downloadAsSource(item: DriveItem): DownloadAsSource | null {
	const base = asDirectoryOrFile(item)

	if (base.type !== "file" || base.data.undecryptable || !canPreview(item)) {
		return null
	}

	switch (previewType(item)) {
		case "docx":
			return "document"
		case "spreadsheet":
			return "spreadsheet"
		default:
			return null
	}
}

// The targets every one of `items` exports to, grouped; empty when any item exports to none (the bulk menus offer
// Download as only when every selected item can be exported, as Extract only when every one is an archive).
export function downloadAsTargets(items: readonly DriveItem[]): DownloadAsTarget[][] {
	const sources: DownloadAsSource[] = []

	for (const item of items) {
		const source = downloadAsSource(item)

		if (source === null) {
			return []
		}

		sources.push(source)
	}

	if (sources.length === 0) {
		return []
	}

	return DOWNLOAD_AS_GROUPS.map(group =>
		group.filter(target => sources.every(source => DOWNLOAD_AS_TARGETS[target].sources.has(source)))
	).filter(group => group.length > 0)
}

export function canDownloadAs(items: readonly DriveItem[]): boolean {
	return downloadAsTargets(items).length > 0
}
