import { useEffect, useRef, useState, type KeyboardEvent, type RefObject } from "react"
import { useTranslation } from "react-i18next"
import { toast } from "sonner"
import {
	TOGGLE_FORMATS,
	newSheetName,
	type LayerKey,
	type SheetCommand,
	type SheetFileKind,
	type SheetFormatPatch,
	type SheetId
} from "@filen/office/core/sheet"
import {
	SheetGrid,
	createSheetGrid,
	editingChanged,
	useSheetGrid,
	type SheetGridHandle,
	type SheetGridSizes
} from "@filen/office/react/sheet"
import "@filen/office/css/sheet.css"
import "@/features/spreadsheet/next/sheetTheme.css"
import { Button } from "@/components/ui/button"
import { ConfirmDialog } from "@/components/dialogs/confirmDialog"
import { InputDialog } from "@/components/dialogs/inputDialog"
import { PasswordBadge } from "@/features/preview/components/docxToolbar"
import { SheetTabs } from "@/features/spreadsheet/components/sheetTabs"
import { useSizeLayer } from "@/features/spreadsheet/hooks/useSizeLayer"
import type { SpreadsheetSnapshot } from "@/features/spreadsheet/hooks/useSpreadsheetEdits"
import { assertSealedUpload } from "@/features/spreadsheet/lib/xlsxEncryption"
import { useSheetClipboard } from "@/features/spreadsheet/next/clipboard/useSheetClipboard"
import { SheetFormulaBar } from "@/features/spreadsheet/next/components/sheetFormulaBar"
import { SheetGridMenu } from "@/features/spreadsheet/next/components/sheetGridMenu"
import type { SheetHandle } from "@/features/spreadsheet/next/hooks/useSheetSession"
import { isTextSheetKind } from "@/features/spreadsheet/next/lib/sheetHosts"
import { readOnlyNoteKey, refusalKey } from "@/features/spreadsheet/next/lib/sheetMessages.logic"
import { SheetToolbar } from "@/features/spreadsheet/next/toolbar/sheetToolbar"
import { resolveTheme, useTheme } from "@/providers/themeProvider"

export interface SheetBodyProps {
	handle: SheetHandle
	kind: SheetFileKind
	alt: string
	// This user may write the file here (their drive, a writable share).
	editable: boolean
	unnamed: boolean
	renamed: boolean
	readOnlyReason: "renamed" | undefined
	// Nothing will ever edit this file here (a public link): the session lets go of what only a save needs.
	neverEditable: boolean
	onDirtyChange: ((dirty: boolean) => void) | undefined
	saveRef: RefObject<(() => Promise<SpreadsheetSnapshot>) | null> | undefined
	// The author recommends read-only: reopens the file for editing.
	editAnyway: (() => void) | undefined
	// Where the sizes kept beside the file live (a text file's, a view-only workbook's).
	layerKey: LayerKey
	// A legacy workbook the user may write beside: converts it to an .xlsx next to it. Resolves once done or failed.
	onSaveAsXlsx: (() => Promise<void>) | undefined
}

// An open session as a grid, with the formula bar, the formatting toolbar, the sheet tabs and the grid's menu. The
// session holds the cells; the grid's model (react/sheet) holds what is on show and talks to the session.
export function SheetBody({
	handle,
	kind,
	alt,
	editable,
	unnamed,
	renamed,
	readOnlyReason,
	neverEditable,
	onDirtyChange,
	saveRef,
	editAnyway,
	layerKey,
	onSaveAsXlsx
}: SheetBodyProps) {
	const { t } = useTranslation(["preview", "common"])
	const [model] = useState(() => createSheetGrid(handle))
	const grid = useSheetGrid(model)
	const { theme } = useTheme()
	// Sizes beside the file, for a session that cannot hold them (the grid decides when): the layer is kept by sheet
	// index, the grid names sheets by id.
	const sizeLayer = useSizeLayer(layerKey)
	const indexOf = (sheet: SheetId): number => grid.sheets.findIndex(entry => entry.id === sheet)
	const sizes: SheetGridSizes = {
		sizes: sheet => sizeLayer.layer.get(indexOf(sheet)),
		update: (sheet, axis, entries) => {
			const index = indexOf(sheet)

			if (index >= 0) {
				sizeLayer.update(index, axis, entries)
			}
		},
		follow: (sheet, shift) => {
			const index = indexOf(sheet)

			if (index >= 0) {
				sizeLayer.follow(index, shift)
			}
		}
	}
	const gridRef = useRef<SheetGridHandle>(null)
	const containerRef = useRef<HTMLDivElement>(null)
	const [renaming, setRenaming] = useState<number | null>(null)
	const [xlsxAsked, setXlsxAsked] = useState(false)
	const [xlsxSaving, setXlsxSaving] = useState(false)
	const text = isTextSheetKind(kind)
	const mayEdit = editable && !unnamed && !renamed && !neverEditable
	const canEdit = mayEdit && grid.writability.state !== "viewOnly" && grid.failed === null
	// A workbook that turns out read-only keeps the toolbar, disabled, so the grid never moves as the verdict lands.
	const toolbarShown = canEdit || (mayEdit && !text)
	const noteKey = readOnlyNoteKey({ editable, readOnlyReason, unnamed, renamed, writability: grid.writability, kind })
	const entryChanged = grid.editing !== null && editingChanged(grid.editing)
	const dirty = grid.state.dirty || entryChanged
	const shownSheets = grid.sheets.filter(sheet => sheet.state === "visible")
	const shownIndex = shownSheets.findIndex(sheet => sheet.id === grid.sheet)
	const activeTab = shownIndex < 0 ? 0 : shownIndex
	const sheetName = grid.sheets.find(sheet => sheet.id === grid.sheet)?.name ?? ""

	useEffect(
		() =>
			handle.events.attach(event => {
				model.handleEvent(event)
			}),
		[handle, model]
	)

	useEffect(() => {
		model.setReadOnly(!mayEdit)
	}, [model, mayEdit])

	useEffect(() => {
		// Nothing here will save it: the session lets go of what only a save needs.
		if (neverEditable || unnamed || grid.writability.state === "viewOnly") {
			handle.session.viewOnly()
		}
	}, [handle, neverEditable, unnamed, grid.writability.state])

	useEffect(() => {
		model.setRefusalHandler(refusal => {
			// One toast per reason: edits queued behind a failing proof are all refused alike.
			toast.error(t(refusalKey(refusal, kind)), { id: refusal.reason })
		})

		return () => {
			model.setRefusalHandler(null)
		}
	}, [model, t, kind])

	useEffect(() => {
		onDirtyChange?.(dirty)
	}, [dirty, onDirtyChange])

	useSheetClipboard({
		model,
		container: containerRef,
		canEdit,
		onRefused: reason => {
			toast.error(t(reason === "tooLarge" ? "previewSpreadsheetCopyTooLarge" : "previewSpreadsheetEditFailed"))
		}
	})

	function run(command: SheetCommand): void {
		void model.intent()
		void model.command(command)
	}

	function format(patch: SheetFormatPatch): void {
		run({ type: "format.apply", sheet: grid.sheet, ranges: grid.selection.ranges, patch })
	}

	// An open entry is committed first; the session saves after every edit already sent. A protected workbook's
	// save comes back sealed, and nothing else leaves it.
	async function save(): Promise<SpreadsheetSnapshot> {
		await model.commitEditing(null)

		const result = await handle.session.save()

		if (!result.ok) {
			throw new Error(`spreadsheet: the save failed (${result.reason}): ${result.message}`)
		}

		if (handle.info.encrypted !== undefined) {
			assertSealedUpload(result.bytes)
		}

		return {
			bytes: result.bytes,
			commit: () => {
				handle.session.markSaved(result.version).then(
					state => {
						model.apply({ state, patches: [], invalidated: [], layouts: [] })
					},
					() => undefined
				)
			}
		}
	}

	useEffect(() => {
		if (saveRef === undefined) {
			return undefined
		}

		saveRef.current = canEdit ? save : null

		return () => {
			saveRef.current = null
		}
	})

	function handleKey(event: KeyboardEvent<HTMLDivElement>): boolean {
		const mod = event.ctrlKey || event.metaKey

		if (!mod || event.altKey || event.shiftKey || !canEdit || text) {
			return false
		}

		const key = event.key.toLowerCase()
		const toggle = TOGGLE_FORMATS.find(entry => entry.shortcutKey === key)

		if (toggle === undefined) {
			return false
		}

		event.preventDefault()

		const { row, col } = grid.selection.active
		const sheet = grid.sheet

		// On or off as the active cell is, as the toolbar button does.
		void model.session.query("cell", { sheet, row, col }).then(
			detail => {
				const style = model.style(detail.style)
				const on = toggle.format === "underline" ? style.underline !== undefined : style[toggle.format] === true

				format(toggle.format === "underline" ? { underline: on ? null : "single" } : { [toggle.format]: !on })
			},
			() => undefined
		)

		return true
	}

	return (
		<div
			ref={containerRef}
			data-preview-surface
			className="flex size-full flex-col"
		>
			{toolbarShown ? (
				<SheetToolbar
					grid={grid}
					disabled={!canEdit}
					formats={!text}
					onFormat={format}
				/>
			) : null}
			<SheetFormulaBar
				grid={grid}
				gridRef={gridRef}
				canEdit={canEdit}
			>
				{grid.failed !== null ? (
					<span
						role="alert"
						className="max-w-1/2 shrink truncate text-xs text-destructive"
					>
						{t("previewSheetFailed")}
					</span>
				) : grid.editing?.invalid === true ? (
					<span
						role="alert"
						className="max-w-1/2 shrink truncate text-xs text-destructive"
					>
						{t("previewSpreadsheetFormulaInvalid")}
					</span>
				) : noteKey === null ? null : (
					<span
						role="status"
						title={t(noteKey)}
						className="max-w-1/2 shrink truncate text-xs text-muted-foreground"
					>
						{t(noteKey)}
					</span>
				)}
				{editAnyway !== undefined && editable && !unnamed && !renamed && !neverEditable ? (
					<Button
						variant="outline"
						size="xs"
						className="shrink-0"
						onClick={editAnyway}
					>
						{t("previewSheetEditAnyway")}
					</Button>
				) : null}
				{handle.info.encrypted === undefined ? null : (
					<PasswordBadge
						integrity={handle.info.encrypted}
						saving={canEdit}
					/>
				)}
				{onSaveAsXlsx === undefined ? null : (
					<Button
						variant="outline"
						size="xs"
						className="shrink-0"
						onClick={() => {
							setXlsxAsked(true)
						}}
					>
						{t("previewSpreadsheetSaveAsXlsx")}
					</Button>
				)}
			</SheetFormulaBar>
			<SheetGridMenu
				grid={grid}
				canEdit={canEdit}
				run={run}
			>
				<SheetGrid
					ref={gridRef}
					grid={model}
					className="min-h-0 flex-1"
					label={sheetName === "" ? alt : `${alt} – ${sheetName}`}
					editorLabel={t("previewSheetCellEditor")}
					sizes={sizes}
					themeKey={resolveTheme(theme)}
					resizeLabel={(axis, px) =>
						t(axis === "cols" ? "previewSpreadsheetWidthPx" : "previewSpreadsheetHeightPx", { size: px })
					}
					onKey={handleKey}
				/>
			</SheetGridMenu>
			{text ? null : (
				<SheetTabs
					sheets={shownSheets}
					active={activeTab}
					onSelect={index => {
						const target = shownSheets[index]

						if (target === undefined || target.id === grid.sheet) {
							return
						}

						void model.commitEditing(null).then(done => {
							if (done) {
								model.setSheet(target.id)
							}
						})
					}}
					onAdd={
						canEdit
							? () => {
									run({
										type: "sheet.add",
										name: newSheetName(
											grid.sheets.map(sheet => sheet.name),
											number => t("previewSpreadsheetNewSheetName", { number })
										)
									})
								}
							: undefined
					}
					onRename={
						canEdit
							? index => {
									setRenaming(shownSheets[index]?.id ?? null)
								}
							: undefined
					}
				/>
			)}
			<InputDialog
				open={canEdit && renaming !== null}
				pending={false}
				title={t("previewSpreadsheetRenameSheet")}
				body={t("previewSpreadsheetRenameSheetBody")}
				label={t("previewSpreadsheetSheetName")}
				initialValue={grid.sheets.find(sheet => sheet.id === renaming)?.name ?? ""}
				submitLabel={t("previewSpreadsheetRenameSheet")}
				validate={value => value.trim().length > 0}
				onOpenChange={open => {
					if (!open) {
						setRenaming(null)
					}
				}}
				onSubmit={value => {
					if (canEdit && renaming !== null) {
						run({ type: "sheet.rename", sheet: renaming, name: value })
					}

					setRenaming(null)
				}}
			/>
			{onSaveAsXlsx === undefined ? null : (
				<ConfirmDialog
					open={xlsxAsked}
					pending={xlsxSaving}
					title={t("previewSpreadsheetSaveAsXlsxTitle")}
					body={t("previewSpreadsheetSaveAsXlsxBody")}
					confirmLabel={t("previewSpreadsheetSaveAsXlsx")}
					cancelLabel={t("common:cancel")}
					onOpenChange={setXlsxAsked}
					onConfirm={() => {
						setXlsxSaving(true)
						void onSaveAsXlsx().finally(() => {
							setXlsxSaving(false)
							setXlsxAsked(false)
						})
					}}
				/>
			)}
		</div>
	)
}
