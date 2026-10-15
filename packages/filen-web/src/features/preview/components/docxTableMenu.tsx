import { Fragment } from "react"
import { useTranslation } from "react-i18next"
import { Table2Icon } from "lucide-react"
import type { DocxCellVAlign, DocxEditorCommand, DocxEditorCommandType, DocxTableAlign } from "@filen/office/core/docx/edit"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import {
	STANDARD_SWATCHES,
	keepFocus,
	type DocxInputChannelProps,
	type DocxSwatch,
	type DocxToolbarCommand,
	type DocxToolbarSlotProps
} from "@/features/preview/components/docxEditor.logic"
import { DocxColorGrid } from "@/features/preview/components/docxFormatControls"
import { DocxRowsColumnsDialog } from "@/features/preview/components/docxFields"
import { Button } from "@/components/ui/button"
import {
	DropdownMenu,
	DropdownMenuCheckboxItem,
	DropdownMenuContent,
	DropdownMenuItem,
	DropdownMenuRadioGroup,
	DropdownMenuRadioItem,
	DropdownMenuSeparator,
	DropdownMenuSub,
	DropdownMenuSubContent,
	DropdownMenuSubTrigger,
	DropdownMenuTrigger
} from "@/components/ui/dropdown-menu"
import type { PreviewKey } from "@/lib/i18n"

// The toolbar's contextual Table group, shown while the caret is in a table: rows and columns, merging and
// splitting cells (a rows × columns dialog), the header row, distributing columns, the table's alignment and
// style, and the selected cells' alignment and shading (the shared colour picker). Column widths are
// dragged on the pages (docxTableBorderDrag.tsx).

export interface DocxTableMenuProps extends DocxToolbarSlotProps {
	// The theme part of the shading picker.
	theme: readonly (readonly DocxSwatch[])[]
}

// Plain commands, by group (a separator between groups).
const GROUPS: readonly (readonly DocxEditorCommandType[])[] = [
	["table.row.insertAbove", "table.row.insertBelow", "table.column.insertLeft", "table.column.insertRight"],
	["table.row.delete", "table.column.delete", "table.delete"],
	["table.cells.merge", "table.cell.split", "table.distributeColumns"]
]

const ALIGNS: readonly { value: DocxTableAlign; label: PreviewKey }[] = [
	{ value: "start", label: "previewDocxTableAlignStart" },
	{ value: "center", label: "previewDocxTableAlignCenter" },
	{ value: "end", label: "previewDocxTableAlignEnd" }
]

const V_ALIGNS: readonly { value: DocxCellVAlign; label: PreviewKey }[] = [
	{ value: "top", label: "previewDocxTableCellTop" },
	{ value: "center", label: "previewDocxTableCellCenter" },
	{ value: "bottom", label: "previewDocxTableCellBottom" }
]

export function DocxTableMenu(props: DocxTableMenuProps) {
	const { t } = useTranslation("preview")
	const { snapshot, readOnly, execute, commands, requestInput } = props
	const table = snapshot?.context.table ?? null
	const has = (type: DocxEditorCommandType): DocxToolbarCommand | undefined => commands.find(command => command.info.type === type)
	const off = (type: DocxEditorCommandType): boolean => snapshot === undefined || readOnly || !snapshot.enabled[type]
	const item = (type: DocxEditorCommandType) => {
		const command = has(type)

		return command === undefined ? null : (
			<DropdownMenuItem
				key={type}
				disabled={off(type)}
				onClick={() => {
					if (type === "table.cell.split") {
						requestInput(type)
					} else {
						execute({ type } as DocxEditorCommand)
					}
				}}
			>
				{t(command.label)}
			</DropdownMenuItem>
		)
	}

	if (table === null) {
		return null
	}

	return (
		<DropdownMenu>
			<DropdownMenuTrigger
				render={
					<Button
						variant="ghost"
						size="sm"
						className="shrink-0 gap-1 px-2"
						onMouseDown={keepFocus}
					>
						<Table2Icon />
						<span className="max-md:sr-only">{t("previewDocxTableGroup")}</span>
					</Button>
				}
			/>
			<DropdownMenuContent data-preview-surface>
				{GROUPS.map((group, index) => (
					<Fragment key={group[0]}>
						{index > 0 ? <DropdownMenuSeparator /> : null}
						{group.map(item)}
					</Fragment>
				))}
				<DropdownMenuSeparator />
				{has("table.row.header") === undefined ? null : (
					<DropdownMenuCheckboxItem
						checked={table.headerRow}
						closeOnClick
						disabled={off("table.row.header")}
						onClick={() => {
							execute({ type: "table.row.header" })
						}}
					>
						{t("previewDocxTableHeaderRow")}
					</DropdownMenuCheckboxItem>
				)}
				{has("table.align") === undefined ? null : (
					<DropdownMenuSub>
						<DropdownMenuSubTrigger disabled={off("table.align")}>{t("previewDocxTableAlign")}</DropdownMenuSubTrigger>
						<DropdownMenuSubContent>
							<DropdownMenuRadioGroup
								value={table.align ?? ""}
								onValueChange={(align: DocxTableAlign) => {
									execute({ type: "table.align", align })
								}}
							>
								{ALIGNS.map(align => (
									<DropdownMenuRadioItem
										key={align.value}
										value={align.value}
										closeOnClick
									>
										{t(align.label)}
									</DropdownMenuRadioItem>
								))}
							</DropdownMenuRadioGroup>
						</DropdownMenuSubContent>
					</DropdownMenuSub>
				)}
				{has("table.cell.vAlign") === undefined ? null : (
					<DropdownMenuSub>
						<DropdownMenuSubTrigger disabled={off("table.cell.vAlign")}>
							{t("previewDocxTableCellAlign")}
						</DropdownMenuSubTrigger>
						<DropdownMenuSubContent>
							<DropdownMenuRadioGroup
								value={table.vAlign === "mixed" ? "" : (table.vAlign ?? "top")}
								onValueChange={(align: DocxCellVAlign) => {
									execute({ type: "table.cell.vAlign", align })
								}}
							>
								{V_ALIGNS.map(align => (
									<DropdownMenuRadioItem
										key={align.value}
										value={align.value}
										closeOnClick
									>
										{t(align.label)}
									</DropdownMenuRadioItem>
								))}
							</DropdownMenuRadioGroup>
						</DropdownMenuSubContent>
					</DropdownMenuSub>
				)}
				{has("table.cell.shading") === undefined ? null : (
					<DropdownMenuSub>
						<DropdownMenuSubTrigger disabled={off("table.cell.shading")}>{t("previewDocxTableShading")}</DropdownMenuSubTrigger>
						<DropdownMenuSubContent className="p-3">
							<DocxColorGrid
								theme={props.theme}
								standard={STANDARD_SWATCHES}
								chosen={table.shading === "mixed" ? undefined : (table.shading ?? "none")}
								none={{ label: t("previewDocxColorNone"), key: "none" }}
								onPick={swatch => {
									const color = swatch?.color

									// A theme colour keeps its reference (Word writes w:themeFill with the fill).
									execute(
										swatch === null
											? { type: "table.cell.shading", fill: null }
											: {
													type: "table.cell.shading",
													fill: swatch.rgb,
													...(color?.kind === "theme"
														? {
																theme: color.theme,
																...(color.tint === undefined ? {} : { tint: color.tint }),
																...(color.shade === undefined ? {} : { shade: color.shade })
															}
														: {})
												}
									)
								}}
							/>
						</DropdownMenuSubContent>
					</DropdownMenuSub>
				)}
				{has("table.style") === undefined || table.styles.length === 0 ? null : (
					<DropdownMenuSub>
						<DropdownMenuSubTrigger disabled={off("table.style")}>{t("previewDocxTableStyle")}</DropdownMenuSubTrigger>
						<DropdownMenuSubContent className="max-h-80 overflow-y-auto">
							<DropdownMenuRadioGroup
								value={table.styleId ?? ""}
								onValueChange={(styleId: string) => {
									execute({ type: "table.style", styleId })
								}}
							>
								{table.styles.map(style => (
									<DropdownMenuRadioItem
										key={style.id}
										value={style.id}
										closeOnClick
									>
										{style.name}
									</DropdownMenuRadioItem>
								))}
							</DropdownMenuRadioGroup>
						</DropdownMenuSubContent>
					</DropdownMenuSub>
				)}
			</DropdownMenuContent>
		</DropdownMenu>
	)
}

export interface DocxTableSplitDialogProps extends DocxInputChannelProps {
	editor: PagedEditor
	// The host's read-only state: nothing is split meanwhile.
	readOnly: boolean
}

// Split cells: the selected cell into rows × columns (table.cell.split's input, from the Table group or the
// context menu).
export function DocxTableSplitDialog({ editor, readOnly, inputRequest, onInputDone }: DocxTableSplitDialogProps) {
	const { t } = useTranslation("preview")

	if (inputRequest?.type !== "table.cell.split") {
		return null
	}

	const nonce = inputRequest.nonce

	return (
		<DocxRowsColumnsDialog
			key={nonce}
			title={t("previewDocxTableSplitTitle")}
			description={t("previewDocxTableSplitBody")}
			submitLabel={t("previewDocxTableSplitAction")}
			initial={{ rows: 1, cols: 2 }}
			accepts={size => size.rows * size.cols > 1}
			// Run here, so a split the table refuses keeps the dialog open and says so.
			onSubmit={size => !readOnly && editor.execute({ type: "table.cell.split", ...size }) === true}
			refusal={t("previewDocxTableSplitRefused")}
			onClose={() => {
				onInputDone(nonce)
			}}
		/>
	)
}
