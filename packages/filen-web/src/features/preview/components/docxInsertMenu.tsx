import { useState, type ReactNode } from "react"
import { useTranslation } from "react-i18next"
import { BookmarkIcon, ImageIcon, LinkIcon, OmegaIcon, PlusIcon, SigmaIcon, TableIcon } from "lucide-react"
import {
	parseDocxEditorCommand,
	type DocxEditorCommand,
	type DocxEditorCommandType,
	type DocxEditorSnapshot
} from "@filen/office/core/docx/edit"
import { cn } from "@filen/shared"
import {
	COMMAND_LABELS,
	commandVariants,
	keepFocus,
	type DocxToolbarCommand,
	type DocxToolbarSlotProps
} from "@/features/preview/components/docxEditor.logic"
import { TABLE_GRID, useCommandMenuFocus } from "@/features/preview/components/docxInsert.logic"
import { Button } from "@/components/ui/button"
import {
	DropdownMenu,
	DropdownMenuContent,
	DropdownMenuItem,
	DropdownMenuSeparator,
	DropdownMenuSub,
	DropdownMenuSubContent,
	DropdownMenuSubTrigger,
	DropdownMenuTrigger
} from "@/components/ui/dropdown-menu"
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip"

// The toolbar's Insert menu: whatever of links, pictures, tables, breaks, notes, fields, form controls,
// equations and symbols is labelled, in Word's order. Items needing input ask for it through the input channel
// (the dialogs in docxInsertDialogs.tsx, the picture sources in docxImageInsert.tsx, the equation popover).

// Commands the menu places itself; others it is given that need no input and act at the caret are listed after.
const PLACED: ReadonlySet<DocxEditorCommandType> = new Set([
	"image.insert",
	"table.insert",
	"link.set",
	"link.remove",
	"bookmark.add",
	"bookmark.rename",
	"bookmark.delete",
	"header.edit",
	"footer.edit",
	"insert.pageNumber",
	"insert.pageCount",
	"note.insertFootnote",
	"note.insertEndnote",
	"note.convertAll",
	"note.options",
	"toc.insert",
	"toc.update",
	"insert.date",
	"insert.field",
	"field.update",
	"field.updateAll",
	"insert.pageBreak",
	"insert.lineBreak",
	"insert.sectionBreak",
	"control.insert",
	"equation.insert",
	"equation.edit",
	"insert.symbol",
	"insert.horizontalLine"
])

function disabledFor(snapshot: DocxEditorSnapshot | undefined, readOnly: boolean, command: DocxToolbarCommand): boolean {
	return snapshot === undefined || (readOnly && command.info.write) || !snapshot.enabled[command.info.type]
}

// Insert ▸ Table: a grid of sizes to click, and the size dialog.
function TableGrid({ disabled, onPick }: { disabled: boolean; onPick: (cols: number, rows: number) => void }) {
	const { t } = useTranslation("preview")
	const [hover, setHover] = useState<{ cols: number; rows: number } | null>(null)

	return (
		<div className="flex flex-col gap-1 p-1">
			<div
				className="grid gap-0.5"
				style={{ gridTemplateColumns: `repeat(${String(TABLE_GRID.cols)}, 1rem)` }}
				onPointerLeave={() => {
					setHover(null)
				}}
			>
				{Array.from({ length: TABLE_GRID.rows * TABLE_GRID.cols }, (_, index) => {
					const cols = (index % TABLE_GRID.cols) + 1
					const rows = Math.floor(index / TABLE_GRID.cols) + 1
					const lit = hover !== null && cols <= hover.cols && rows <= hover.rows

					return (
						<button
							key={index}
							type="button"
							disabled={disabled}
							aria-label={t("previewDocxInsertTableGrid", { cols, rows })}
							className={cn("size-4 rounded-sm border border-border", lit ? "bg-primary/30" : "bg-background")}
							onPointerEnter={() => {
								setHover({ cols, rows })
							}}
							onFocus={() => {
								setHover({ cols, rows })
							}}
							onClick={() => {
								onPick(cols, rows)
							}}
						/>
					)
				})}
			</div>
			<p
				aria-hidden="true"
				className="h-5 text-center text-xs text-muted-foreground"
			>
				{hover === null ? "" : `${String(hover.cols)} × ${String(hover.rows)}`}
			</p>
		</div>
	)
}

export function DocxInsertMenu(props: DocxToolbarSlotProps) {
	const { t } = useTranslation("preview")
	const { commands, snapshot, readOnly, execute, requestInput } = props
	const [open, setOpen] = useState(false)
	const focus = useCommandMenuFocus()
	const runCommand = (command: DocxEditorCommand): void => {
		execute(command)
		focus.ran()
	}
	const byType = new Map(commands.map(command => [command.info.type, command]))
	const command = (type: DocxEditorCommandType): DocxToolbarCommand | undefined => byType.get(type)
	const off = (type: DocxEditorCommandType): boolean => {
		const found = command(type)

		return found === undefined || disabledFor(snapshot, readOnly, found)
	}
	const label = (type: DocxEditorCommandType): string => {
		const key = COMMAND_LABELS[type]

		return key === undefined ? type : t(key)
	}
	// An item running a command without arguments.
	const runItem = (type: DocxEditorCommandType, icon?: ReactNode): ReactNode =>
		command(type) === undefined ? null : (
			<DropdownMenuItem
				key={type}
				disabled={off(type)}
				onClick={() => {
					const parsed = parseDocxEditorCommand({ type })

					if (parsed !== undefined) {
						runCommand(parsed)
					}
				}}
			>
				{icon}
				{label(type)}
			</DropdownMenuItem>
		)
	// An item asking for the command's input.
	const inputItem = (type: DocxEditorCommandType, icon?: ReactNode): ReactNode =>
		command(type) === undefined ? null : (
			<DropdownMenuItem
				key={type}
				disabled={off(type)}
				onClick={() => {
					requestInput(type)
				}}
			>
				{icon}
				{label(type)}
			</DropdownMenuItem>
		)
	// A submenu of the forms a command is offered in.
	const variantsItem = (type: DocxEditorCommandType): ReactNode => {
		const variants = commandVariants(type)

		if (command(type) === undefined || variants === undefined) {
			return null
		}

		return (
			<DropdownMenuSub key={type}>
				<DropdownMenuSubTrigger disabled={off(type)}>{label(type)}</DropdownMenuSubTrigger>
				<DropdownMenuSubContent>
					{variants.map(variant => (
						<DropdownMenuItem
							key={variant.label}
							onClick={() => {
								if ("command" in variant) {
									runCommand(variant.command)
								} else {
									requestInput(type, variant.input)
								}
							}}
						>
							{t(variant.label)}
						</DropdownMenuItem>
					))}
				</DropdownMenuSubContent>
			</DropdownMenuSub>
		)
	}
	const hasAny = (...types: DocxEditorCommandType[]): boolean => types.some(type => command(type) !== undefined)
	const allOff = (...types: DocxEditorCommandType[]): boolean => types.every(type => off(type))
	const rest = commands.filter(
		entry => !PLACED.has(entry.info.type) && entry.info.input === undefined && entry.info.context === undefined
	)
	const link = snapshot?.context.link ?? null

	return (
		<DropdownMenu
			open={open}
			onOpenChange={setOpen}
		>
			<Tooltip>
				<TooltipTrigger
					render={
						<DropdownMenuTrigger
							render={
								<Button
									variant="ghost"
									size="sm"
									aria-label={t("previewDocxInsertMenu")}
									data-docx-insert-menu
									className="shrink-0 gap-1 px-2"
									onMouseDown={keepFocus}
								>
									<PlusIcon />
									<span className="max-md:hidden">{t("previewDocxInsertMenu")}</span>
								</Button>
							}
						/>
					}
				/>
				<TooltipContent>{t("previewDocxInsertMenu")}</TooltipContent>
			</Tooltip>
			<DropdownMenuContent
				align="start"
				finalFocus={focus.finalFocus}
				data-preview-surface
			>
				{command("image.insert") === undefined ? null : (
					<DropdownMenuSub>
						<DropdownMenuSubTrigger disabled={off("image.insert")}>
							<ImageIcon />
							{label("image.insert")}
						</DropdownMenuSubTrigger>
						<DropdownMenuSubContent>
							<DropdownMenuItem
								onClick={() => {
									requestInput("image.insert", "device")
								}}
							>
								{t("previewDocxPictureFromDevice")}
							</DropdownMenuItem>
							<DropdownMenuItem
								onClick={() => {
									requestInput("image.insert", "drive")
								}}
							>
								{t("previewDocxPictureFromDrive")}
							</DropdownMenuItem>
						</DropdownMenuSubContent>
					</DropdownMenuSub>
				)}
				{command("table.insert") === undefined ? null : (
					<DropdownMenuSub>
						<DropdownMenuSubTrigger disabled={off("table.insert")}>
							<TableIcon />
							{t("previewDocxInsertTable")}
						</DropdownMenuSubTrigger>
						<DropdownMenuSubContent>
							<TableGrid
								disabled={off("table.insert")}
								onPick={(cols, rows) => {
									setOpen(false)
									runCommand({ type: "table.insert", cols, rows })
								}}
							/>
							<DropdownMenuSeparator />
							<DropdownMenuItem
								onClick={() => {
									requestInput("table.insert")
								}}
							>
								{t("previewDocxInsertTableCustom")}
							</DropdownMenuItem>
						</DropdownMenuSubContent>
					</DropdownMenuSub>
				)}
				{inputItem("link.set", <LinkIcon />)}
				{link === null ? null : runItem("link.remove")}
				{inputItem("bookmark.add", <BookmarkIcon />)}
				{hasAny("header.edit", "footer.edit") ? <DropdownMenuSeparator /> : null}
				{runItem("header.edit")}
				{runItem("footer.edit")}
				{hasAny("insert.pageNumber", "insert.pageCount") ? (
					<DropdownMenuSub>
						<DropdownMenuSubTrigger disabled={allOff("insert.pageNumber", "insert.pageCount")}>
							{t("previewDocxInsertPageNumbers")}
						</DropdownMenuSubTrigger>
						<DropdownMenuSubContent>
							{runItem("insert.pageNumber")}
							{runItem("insert.pageCount")}
						</DropdownMenuSubContent>
					</DropdownMenuSub>
				) : null}
				{hasAny("note.insertFootnote", "note.insertEndnote", "note.convertAll", "note.options") ? (
					<DropdownMenuSub>
						<DropdownMenuSubTrigger
							disabled={allOff("note.insertFootnote", "note.insertEndnote", "note.convertAll", "note.options")}
						>
							{t("previewDocxInsertNotes")}
						</DropdownMenuSubTrigger>
						<DropdownMenuSubContent>
							{runItem("note.insertFootnote")}
							{runItem("note.insertEndnote")}
							{command("note.convertAll") === undefined ? null : (
								<>
									<DropdownMenuSeparator />
									<DropdownMenuItem
										disabled={off("note.convertAll")}
										onClick={() => {
											runCommand({ type: "note.convertAll", to: "endnote" })
										}}
									>
										{t("previewDocxNoteConvertAllToEndnotes")}
									</DropdownMenuItem>
									<DropdownMenuItem
										disabled={off("note.convertAll")}
										onClick={() => {
											runCommand({ type: "note.convertAll", to: "footnote" })
										}}
									>
										{t("previewDocxNoteConvertAllToFootnotes")}
									</DropdownMenuItem>
								</>
							)}
							{command("note.options") === undefined ? null : <DropdownMenuSeparator />}
							{inputItem("note.options")}
						</DropdownMenuSubContent>
					</DropdownMenuSub>
				) : null}
				{hasAny("toc.insert", "insert.date", "insert.field", "field.updateAll") ? <DropdownMenuSeparator /> : null}
				{inputItem("toc.insert")}
				{inputItem("insert.date")}
				{inputItem("insert.field")}
				{runItem("field.updateAll")}
				{hasAny("insert.pageBreak", "insert.lineBreak", "insert.sectionBreak") ? (
					<>
						<DropdownMenuSeparator />
						<DropdownMenuSub>
							<DropdownMenuSubTrigger disabled={allOff("insert.pageBreak", "insert.lineBreak", "insert.sectionBreak")}>
								{t("previewDocxInsertBreaks")}
							</DropdownMenuSubTrigger>
							<DropdownMenuSubContent>
								{runItem("insert.pageBreak")}
								{runItem("insert.lineBreak")}
								{variantsItem("insert.sectionBreak")}
							</DropdownMenuSubContent>
						</DropdownMenuSub>
					</>
				) : null}
				{variantsItem("control.insert")}
				{hasAny("equation.insert", "insert.symbol", "insert.horizontalLine") ? <DropdownMenuSeparator /> : null}
				{inputItem("equation.insert", <SigmaIcon />)}
				{inputItem("insert.symbol", <OmegaIcon />)}
				{runItem("insert.horizontalLine")}
				{rest.length === 0 ? null : <DropdownMenuSeparator />}
				{rest.map(entry => runItem(entry.info.type))}
			</DropdownMenuContent>
		</DropdownMenu>
	)
}
