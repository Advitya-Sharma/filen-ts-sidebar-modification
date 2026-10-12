import { useState, type ReactNode } from "react"
import { useTranslation } from "react-i18next"
import {
	AlignCenterIcon,
	AlignJustifyIcon,
	AlignLeftIcon,
	AlignRightIcon,
	BoldIcon,
	ItalicIcon,
	ListIcon,
	ListIndentDecreaseIcon,
	ListIndentIncreaseIcon,
	ListOrderedIcon,
	LockIcon,
	MoreHorizontalIcon,
	Redo2Icon,
	RemoveFormattingIcon,
	SearchIcon,
	StrikethroughIcon,
	UnderlineIcon,
	Undo2Icon,
	type LucideIcon
} from "lucide-react"
import {
	DOCX_PARAGRAPH_STYLE_MENU,
	parseDocxEditorCommand,
	type DocxEditorCommand,
	type DocxEditorCommandInfo,
	type DocxEditorSnapshot,
	type DocxParagraphStyle,
	type LockReason
} from "@filen/office/core/docx/edit"
import { DOCX_MARKUP_MODES, type DocxMarkupMode } from "@filen/office/dom/docx"
import { cn } from "@filen/shared"
import {
	commandActive,
	lockReasonKey,
	rowCommands,
	shortcutKeys,
	slotCommands,
	styleLabel,
	toolbarFit,
	type DocxToolbarCommand,
	type DocxToolbarSlotProps
} from "@/features/preview/components/docxEditor.logic"
import { DocxFormatControls } from "@/features/preview/components/docxFormatControls"
import { DocxInsertMenu } from "@/features/preview/components/docxInsertMenu"
import { DocxReflowToggle } from "@/features/preview/components/docxReflowSurface"
import { DocxReviewMenu } from "@/features/preview/components/docxReviewMenu"
import { DocxTableMenu } from "@/features/preview/components/docxTableMenu"
import { MARKUP_LABELS, type DocxFindControls, type DocxViewOnlyReason } from "@/features/preview/components/docxViewer.logic"
import { DocxFindField, MarkupMenu, ViewOnlyBadge } from "@/features/preview/components/docxToolbar"
import { Badge } from "@/components/ui/badge"
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
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover"
import { Select, SelectContent, SelectGroup, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select"
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip"
import { TooltipIconButton } from "@/components/ui/tooltipIconButton"
import { Kbd, KbdKeys } from "@/lib/keymap/kbd"
import { useElementSize } from "@/lib/useElementSize"

// The DOCX editor's toolbar: one compact row in the spreadsheet toolbar's style (undo and redo, paragraph
// style, character formatting, alignment, lists), groups that do not fit moved into the ⋯ menu with the
// editor's options (find and replace, markup, zoom, autosave). Mod+F swaps the row for the find field.
// Everything shown comes from the editing session's snapshot; everything done is one of its commands. A
// feature's commands appear through its slot (insert, format, table, review, view) once they are labelled.

const ICONS: Readonly<Record<string, LucideIcon>> = {
	bold: BoldIcon,
	italic: ItalicIcon,
	underline: UnderlineIcon,
	strikethrough: StrikethroughIcon,
	"clear-formatting": RemoveFormattingIcon,
	"align-start": AlignLeftIcon,
	"align-center": AlignCenterIcon,
	"align-end": AlignRightIcon,
	"align-justify": AlignJustifyIcon,
	"list-bullet": ListIcon,
	"list-numbered": ListOrderedIcon,
	indent: ListIndentIncreaseIcon,
	outdent: ListIndentDecreaseIcon,
	undo: Undo2Icon,
	redo: Redo2Icon
}

const HISTORY = rowCommands("history")
// Character formatting without "clear", which has a place of its own (it goes to the menu first).
const FORMAT = rowCommands("format").filter(command => command.info.type !== "format.clear")
const CLEAR = rowCommands("format").filter(command => command.info.type === "format.clear")
const ALIGN = rowCommands("align")
const LISTS = rowCommands("list")
const INSERT_SLOT = slotCommands("insert")
const FORMAT_SLOT = slotCommands("format")
const TABLE_SLOT = slotCommands("table")
const REVIEW_SLOT = slotCommands("review")
const VIEW_SLOT = slotCommands("view")

export interface DocxEditToolbarProps {
	snapshot: DocxEditorSnapshot | undefined
	// The host's read-only state on top of the session's (a save writing, a file that cannot be saved).
	readOnly: boolean
	apple: boolean
	execute: (command: DocxEditorCommand) => void
	viewOnly: readonly DocxViewOnlyReason[]
	find: { open: boolean; replace: boolean; focus: number; controls: DocxFindControls }
	onFindOpen: (replace: boolean) => void
	onFindClosed: () => void
	markup: DocxMarkupMode | null
	onMarkupChange: (mode: DocxMarkupMode) => void
	zoom: number
	onZoom: (change: "in" | "out" | "fit") => void
	autosave: boolean
	onAutosaveChange: (autosave: boolean) => void
	// The reflowed (mobile) view instead of pages.
	reflow: boolean
	onReflowChange: (reflow: boolean) => void
}

function Divider() {
	return (
		<span
			aria-hidden="true"
			className="mx-1 h-5 w-px shrink-0 bg-border"
		/>
	)
}

// A command without arguments (every one but style.set, which the style picker sends).
function run(info: DocxEditorCommandInfo, execute: (command: DocxEditorCommand) => void): void {
	const command = parseDocxEditorCommand({ type: info.type })

	if (command !== undefined) {
		execute(command)
	}
}

// Keeps focus where it was (the document): a pressed toolbar button never takes it.
function keepFocus(event: { preventDefault: () => void }): void {
	event.preventDefault()
}

// Whether a toolbar command is off: no session yet, a write while read-only, or the session says so.
function commandDisabled(info: DocxEditorCommandInfo, snapshot: DocxEditorSnapshot | undefined, readOnly: boolean): boolean {
	return snapshot === undefined || (readOnly && info.write) || !snapshot.enabled[info.type]
}

// Takes what it shows as plain values (never the snapshot, new on every keystroke), so a keystroke that
// changes neither re-renders nothing below it.
function CommandButton({
	command,
	active,
	disabled,
	apple,
	execute
}: {
	command: DocxToolbarCommand
	active: boolean
	disabled: boolean
	apple: boolean
	execute: (command: DocxEditorCommand) => void
}) {
	const { t } = useTranslation("preview")
	const { info, label } = command
	const Icon = ICONS[info.icon ?? ""] ?? BoldIcon
	const toggle = info.group === "format" && info.type !== "format.clear"

	return (
		<TooltipIconButton
			label={t(label)}
			disabled={disabled}
			pressed={toggle || info.group === "align" || info.type === "list.bullet" || info.type === "list.numbered" ? active : undefined}
			className={cn(active && "bg-accent text-foreground")}
			shortcut={info.shortcut === undefined ? undefined : <KbdKeys keys={shortcutKeys(info.shortcut, apple)} />}
			onMouseDown={keepFocus}
			onClick={() => {
				run(info, execute)
			}}
		>
			<Icon />
		</TooltipIconButton>
	)
}

const STYLE_MIXED = "mixed"
const STYLE_CUSTOM = "custom"

function StylePicker({
	style,
	styleName,
	disabled,
	execute
}: {
	style: DocxEditorSnapshot["paragraph"]["style"]
	styleName: string
	disabled: boolean
	execute: (command: DocxEditorCommand) => void
}) {
	const { t } = useTranslation("preview")
	const value = style ?? STYLE_MIXED
	const labelOf = (option: DocxParagraphStyle): string => {
		const label = styleLabel(option)

		return label.key === "previewDocxStyleHeading" ? t(label.key, { level: label.level }) : t(label.key)
	}
	const items = [
		...DOCX_PARAGRAPH_STYLE_MENU.map(option => ({ value: option.style, label: labelOf(option.style) })),
		{ value: STYLE_CUSTOM, label: styleName },
		{ value: STYLE_MIXED, label: t("previewDocxStyleMixed") }
	]

	return (
		<Select
			items={items}
			value={value}
			disabled={disabled}
			onValueChange={next => {
				const chosen = DOCX_PARAGRAPH_STYLE_MENU.find(option => option.style === next)

				if (chosen !== undefined) {
					execute({ type: "style.set", style: chosen.style })
				}
			}}
		>
			<SelectTrigger
				size="sm"
				aria-label={t("previewDocxParagraphStyle")}
				className="w-36 shrink-0"
			>
				<SelectValue />
			</SelectTrigger>
			<SelectContent data-preview-surface>
				<SelectGroup>
					{DOCX_PARAGRAPH_STYLE_MENU.map(option => (
						<SelectItem
							key={option.style}
							value={option.style}
						>
							{labelOf(option.style)}
						</SelectItem>
					))}
					{value === STYLE_CUSTOM ? <SelectItem value={STYLE_CUSTOM}>{styleName}</SelectItem> : null}
					{value === STYLE_MIXED ? <SelectItem value={STYLE_MIXED}>{t("previewDocxStyleMixed")}</SelectItem> : null}
				</SelectGroup>
			</SelectContent>
		</Select>
	)
}

function LockedBadge({ reason }: { reason: LockReason }) {
	const { t } = useTranslation("preview")

	return (
		<Popover>
			<Badge
				variant="outline"
				className="cursor-pointer text-muted-foreground hover:bg-muted"
				render={<PopoverTrigger aria-label={t("previewDocxLocked")} />}
			>
				<LockIcon data-icon="inline-start" />
				<span className="max-sm:hidden">{t("previewDocxLocked")}</span>
			</Badge>
			<PopoverContent
				align="start"
				className="gap-2"
			>
				<p className="font-medium">{t("previewDocxLockedTitle")}</p>
				<p className="text-muted-foreground">{t(lockReasonKey(reason))}</p>
			</PopoverContent>
		</Popover>
	)
}

// Menu items for the commands of a group the row has no room for.
function OverflowCommands({
	commands,
	snapshot,
	readOnly,
	execute
}: {
	commands: readonly DocxToolbarCommand[]
	snapshot: DocxEditorSnapshot | undefined
	readOnly: boolean
	execute: (command: DocxEditorCommand) => void
}) {
	const { t } = useTranslation("preview")

	return (
		<>
			{commands.map(({ info, label }) => {
				const Icon = ICONS[info.icon ?? ""] ?? BoldIcon
				const disabled = commandDisabled(info, snapshot, readOnly)
				const active = snapshot !== undefined && commandActive(snapshot, info.type)

				return active ? (
					<DropdownMenuCheckboxItem
						key={info.type}
						checked
						disabled={disabled}
						onClick={() => {
							run(info, execute)
						}}
					>
						<Icon />
						{t(label)}
					</DropdownMenuCheckboxItem>
				) : (
					<DropdownMenuItem
						key={info.type}
						disabled={disabled}
						onClick={() => {
							run(info, execute)
						}}
					>
						<Icon />
						{t(label)}
					</DropdownMenuItem>
				)
			})}
		</>
	)
}

export function DocxEditToolbar(props: DocxEditToolbarProps) {
	const { t } = useTranslation("preview")
	const { snapshot, readOnly, apple, execute, find } = props
	const [row, setRow] = useState<HTMLDivElement | null>(null)
	const { width } = useElementSize(row)
	const locked = snapshot?.selection.locked ?? null
	// Badges take room from the groups.
	const fit = toolbarFit(width, (props.viewOnly.length > 0 ? 100 : 0) + (locked === null ? 0 : 120) + (props.markup === null ? 0 : 40))
	const percent = Math.round(props.zoom * 100)
	const commandProps = { snapshot, readOnly, execute }
	const slotProps = (commands: readonly DocxToolbarCommand[]): DocxToolbarSlotProps => ({ commands, snapshot, readOnly, apple, execute })
	const button = (command: DocxToolbarCommand): ReactNode => (
		<CommandButton
			key={command.info.type}
			command={command}
			active={snapshot !== undefined && commandActive(snapshot, command.info.type)}
			disabled={commandDisabled(command.info, snapshot, readOnly)}
			apple={apple}
			execute={execute}
		/>
	)
	let badges: ReactNode = null

	if (props.viewOnly.length > 0 || locked !== null) {
		badges = (
			<div className="flex shrink-0 items-center gap-1">
				{props.viewOnly.length > 0 ? <ViewOnlyBadge reasons={props.viewOnly} /> : null}
				{locked === null ? null : <LockedBadge reason={locked} />}
			</div>
		)
	}

	return (
		<div
			ref={setRow}
			role="toolbar"
			aria-label={t("previewDocxEditToolbar")}
			className="flex h-10 shrink-0 items-center gap-0.5 overflow-hidden border-b border-border px-2"
		>
			{find.open ? (
				<>
					{badges}
					<div className="flex min-w-0 flex-1 justify-center">
						<DocxFindField
							key={find.replace ? "replace" : "find"}
							find={find.controls}
							focus={find.focus}
							onClosed={props.onFindClosed}
							replaceOpen={find.replace}
						/>
					</div>
				</>
			) : (
				<>
					{HISTORY.map(button)}
					{fit.style ? (
						<>
							<Divider />
							<StylePicker
								style={snapshot?.paragraph.style ?? null}
								styleName={snapshot?.paragraph.styleName ?? ""}
								disabled={snapshot === undefined || readOnly || !snapshot.enabled["style.set"]}
								execute={execute}
							/>
						</>
					) : null}
					<Divider />
					{FORMAT.map(button)}
					{fit.clear ? CLEAR.map(button) : null}
					{fit.align ? (
						<>
							<Divider />
							{ALIGN.map(button)}
						</>
					) : null}
					{fit.lists ? (
						<>
							<Divider />
							{LISTS.map(button)}
						</>
					) : null}
					{FORMAT_SLOT.length > 0 ? <DocxFormatControls {...slotProps(FORMAT_SLOT)} /> : null}
					{INSERT_SLOT.length > 0 ? <DocxInsertMenu {...slotProps(INSERT_SLOT)} /> : null}
					{TABLE_SLOT.length > 0 && snapshot?.context.table != null ? <DocxTableMenu {...slotProps(TABLE_SLOT)} /> : null}
					<div className="min-w-0 flex-1" />
					{badges}
					{REVIEW_SLOT.length > 0 ? (
						<DocxReviewMenu
							{...slotProps(REVIEW_SLOT)}
							markup={props.markup}
							onMarkupChange={props.onMarkupChange}
						/>
					) : null}
					{VIEW_SLOT.length > 0 ? (
						<DocxReflowToggle
							{...slotProps(VIEW_SLOT)}
							reflow={props.reflow}
							onReflowChange={props.onReflowChange}
						/>
					) : null}
					{props.markup === null || REVIEW_SLOT.length > 0 ? null : (
						<MarkupMenu
							mode={props.markup}
							onChange={props.onMarkupChange}
						/>
					)}
					<TooltipIconButton
						label={t("previewDocxFindAction")}
						shortcut={<Kbd action="editor.find" />}
						onMouseDown={keepFocus}
						onClick={() => {
							props.onFindOpen(false)
						}}
					>
						<SearchIcon />
					</TooltipIconButton>
				</>
			)}
			<DropdownMenu>
				<Tooltip>
					<TooltipTrigger
						render={
							<DropdownMenuTrigger
								render={
									<Button
										variant="ghost"
										size="icon-sm"
										aria-label={t("previewDocxMoreTools")}
										className="shrink-0"
									>
										<MoreHorizontalIcon />
									</Button>
								}
							/>
						}
					/>
					<TooltipContent>{t("previewDocxMoreTools")}</TooltipContent>
				</Tooltip>
				<DropdownMenuContent
					align="end"
					data-preview-surface
				>
					{!find.open && !fit.style ? (
						<DropdownMenuSub>
							<DropdownMenuSubTrigger disabled={snapshot === undefined || readOnly || !snapshot.enabled["style.set"]}>
								{t("previewDocxParagraphStyle")}
							</DropdownMenuSubTrigger>
							<DropdownMenuSubContent>
								<DropdownMenuRadioGroup
									value={snapshot?.paragraph.style ?? ""}
									onValueChange={(style: DocxParagraphStyle) => {
										execute({ type: "style.set", style })
									}}
								>
									{DOCX_PARAGRAPH_STYLE_MENU.map(option => {
										const label = styleLabel(option.style)

										return (
											<DropdownMenuRadioItem
												key={option.style}
												value={option.style}
												closeOnClick
											>
												{label.key === "previewDocxStyleHeading"
													? t(label.key, { level: label.level })
													: t(label.key)}
											</DropdownMenuRadioItem>
										)
									})}
								</DropdownMenuRadioGroup>
							</DropdownMenuSubContent>
						</DropdownMenuSub>
					) : null}
					{!find.open && !fit.clear ? (
						<OverflowCommands
							commands={CLEAR}
							{...commandProps}
						/>
					) : null}
					{!find.open && !fit.align ? (
						<OverflowCommands
							commands={ALIGN}
							{...commandProps}
						/>
					) : null}
					{!find.open && !fit.lists ? (
						<OverflowCommands
							commands={LISTS}
							{...commandProps}
						/>
					) : null}
					{!find.open && !(fit.style && fit.clear && fit.align && fit.lists) ? <DropdownMenuSeparator /> : null}
					<DropdownMenuItem
						onClick={() => {
							props.onFindOpen(true)
						}}
					>
						{t("previewDocxFindReplaceAction")}
					</DropdownMenuItem>
					{props.markup === null || !find.open ? null : (
						<DropdownMenuSub>
							<DropdownMenuSubTrigger>{t("previewDocxMarkup")}</DropdownMenuSubTrigger>
							<DropdownMenuSubContent>
								<DropdownMenuRadioGroup
									value={props.markup}
									onValueChange={(mode: DocxMarkupMode) => {
										props.onMarkupChange(mode)
									}}
								>
									{DOCX_MARKUP_MODES.map(info => (
										<DropdownMenuRadioItem
											key={info.mode}
											value={info.mode}
											closeOnClick
										>
											{t(MARKUP_LABELS[info.mode])}
										</DropdownMenuRadioItem>
									))}
								</DropdownMenuRadioGroup>
							</DropdownMenuSubContent>
						</DropdownMenuSub>
					)}
					<DropdownMenuSub>
						<DropdownMenuSubTrigger>{t("previewDocxZoom")}</DropdownMenuSubTrigger>
						<DropdownMenuSubContent>
							<DropdownMenuItem
								closeOnClick={false}
								onClick={() => {
									props.onZoom("in")
								}}
							>
								{t("previewDocxZoomInAction")}
							</DropdownMenuItem>
							<DropdownMenuItem
								closeOnClick={false}
								onClick={() => {
									props.onZoom("out")
								}}
							>
								{t("previewDocxZoomOutAction")}
							</DropdownMenuItem>
							<DropdownMenuItem
								onClick={() => {
									props.onZoom("fit")
								}}
							>
								{t("previewDocxFitWidth")}
							</DropdownMenuItem>
						</DropdownMenuSubContent>
					</DropdownMenuSub>
					<DropdownMenuSeparator />
					<DropdownMenuCheckboxItem
						checked={props.autosave}
						onCheckedChange={checked => {
							props.onAutosaveChange(checked)
						}}
					>
						{t("previewDocxAutosave")}
					</DropdownMenuCheckboxItem>
				</DropdownMenuContent>
			</DropdownMenu>
			<Tooltip>
				<TooltipTrigger
					render={
						<Button
							variant="ghost"
							size="xs"
							aria-label={t("previewDocxFitWidthAction", { percent })}
							className="w-12 shrink-0 px-0 text-muted-foreground tabular-nums max-sm:hidden"
							onMouseDown={keepFocus}
							onClick={() => {
								props.onZoom("fit")
							}}
						>
							{t("previewDocxZoomIndicator", { percent })}
						</Button>
					}
				/>
				<TooltipContent>{t("previewDocxFitWidth")}</TooltipContent>
			</Tooltip>
		</div>
	)
}
