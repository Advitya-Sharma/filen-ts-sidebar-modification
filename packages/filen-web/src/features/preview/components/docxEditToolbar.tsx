import { useState, useSyncExternalStore, type ReactNode } from "react"
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
	parseDocxEditorCommand,
	type DocxEditorCommand,
	type DocxEditorCommandInfo,
	type DocxEditorSnapshot,
	type DocxFontChoice,
	type DocxParagraphStyle,
	type LockReason
} from "@filen/office/core/docx/edit"
import { DOCX_MARKUP_MODES, type DocxMarkupMode } from "@filen/office/dom/docx"
import { cn } from "@filen/shared"
import {
	COMMAND_LABELS,
	LINE_SPACINGS,
	commandActive,
	featureCommands,
	keepFocus,
	lockReasonKey,
	rowCommands,
	shortcutKeys,
	slotCommands,
	styleLabel,
	toolbarFit,
	drawingItems,
	single,
	KEPT_FORMAT_BUTTONS,
	slotsWidth,
	type DocxInputRequest,
	type DocxPainterState,
	type DocxRequestInput,
	type DocxSurface,
	type DocxSwatch,
	type DocxToolbarCommand,
	type DocxToolbarSlotProps
} from "@/features/preview/components/docxEditor.logic"
import { DocxDrawingMenu, DocxDrawingOverflow } from "@/features/preview/components/docxDrawingMenu"
import { DocxFormatControls } from "@/features/preview/components/docxFormatControls"
import { DocxInsertMenu } from "@/features/preview/components/docxInsertMenu"
import { DocxReflowMenuItem } from "@/features/preview/components/docxReflowSurface"
import { reflowCommands } from "@/features/preview/components/docxReflow.logic"
import { DocxReviewMenu } from "@/features/preview/components/docxReviewMenu"
import { DocxTableMenu } from "@/features/preview/components/docxTableMenu"
import { MARKUP_LABELS, type DocxFindControls, type DocxViewOnlyReason } from "@/features/preview/components/docxViewer.logic"
import { DocxFindField, MarkupMenu, PasswordBadge, ViewOnlyBadge, type PasswordIntegrity } from "@/features/preview/components/docxToolbar"
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
// style, character formatting, font and colours, alignment, lists), groups that do not fit moved into the ⋯
// menu with the editor's options (find and replace, line spacing, the Paragraph and Page setup dialogs, field
// codes, markup, zoom, autosave). Mod+F swaps the row for the find field.
// Everything shown comes from the editing session's snapshot; everything done is one of its commands. A
// feature's commands appear through its slot (insert, format, table, picture, review, view) once they are labelled.

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
// The contextual Picture group's.
const DRAWING_GROUP = featureCommands("drawings")
// The ⋯ menu's paragraph spacing and indents (the Paragraph dialog takes paragraph.indent's input).
const LINE_SPACING = COMMAND_LABELS["paragraph.lineSpacing"] !== undefined
const PARAGRAPH_DIALOG = COMMAND_LABELS["paragraph.indent"] !== undefined || COMMAND_LABELS["paragraph.spacing"] !== undefined
const PAGE_SETUP = COMMAND_LABELS["page.setup"] !== undefined

export interface DocxEditToolbarProps {
	// The File menu (Download as), first in the row; absent when it offers nothing.
	fileMenu?: ReactNode
	snapshot: DocxEditorSnapshot | undefined
	// The host's read-only state on top of the session's (a save writing, a file that cannot be saved).
	readOnly: boolean
	apple: boolean
	execute: (command: DocxEditorCommand) => void
	viewOnly: readonly DocxViewOnlyReason[]
	// Opened with its password: its integrity, and whether saves keep the password now (the session can save).
	encrypted?: { integrity: PasswordIntegrity; saving: boolean } | undefined
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
	surface: DocxSurface
	// The paragraph styles the picker offers (offeredParagraphStyles), the document's fonts and the theme part
	// of the colour pickers: the same for the editor's lifetime.
	styles: readonly DocxParagraphStyle[]
	fonts: readonly DocxFontChoice[]
	theme: readonly (readonly DocxSwatch[])[]
	painter: DocxPainterState
	onPainterCancel: () => void
	// Every field shows its code (view.fieldCodes).
	fieldCodes: boolean
	// The input request channel the slots take part in (docxEditor.logic.ts DocxInputRequest).
	inputRequest: DocxInputRequest | null
	requestInput: DocxRequestInput
	onInputDone: (nonce: number) => void
}

// The slot menus show their names from this window width (their labels' `md` breakpoint) on.
const SLOT_LABELS_QUERY = "(min-width: 48rem)"

function subscribeSlotLabels(change: () => void): () => void {
	const query = window.matchMedia(SLOT_LABELS_QUERY)

	query.addEventListener("change", change)

	return () => {
		query.removeEventListener("change", change)
	}
}

function slotLabelsShown(): boolean {
	return window.matchMedia(SLOT_LABELS_QUERY).matches
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
	styles,
	style,
	styleName,
	disabled,
	execute
}: {
	styles: readonly DocxParagraphStyle[]
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
	// A style the document's protection locks shows as the selection's own, not as an offer.
	const offered = styles.includes(value as DocxParagraphStyle)
	const shown = offered || value === STYLE_MIXED ? value : STYLE_CUSTOM
	const items = [
		...styles.map(option => ({ value: option, label: labelOf(option) })),
		{ value: STYLE_CUSTOM, label: value === STYLE_CUSTOM ? styleName : labelOf(value as DocxParagraphStyle) },
		{ value: STYLE_MIXED, label: t("previewDocxStyleMixed") }
	]

	return (
		<Select
			items={items}
			value={shown}
			disabled={disabled}
			onValueChange={next => {
				const chosen = styles.find(option => option === next)

				if (chosen !== undefined) {
					execute({ type: "style.set", style: chosen })
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
					{styles.map(option => (
						<SelectItem
							key={option}
							value={option}
						>
							{labelOf(option)}
						</SelectItem>
					))}
					{shown === STYLE_CUSTOM ? (
						<SelectItem value={STYLE_CUSTOM}>{items.find(item => item.value === STYLE_CUSTOM)?.label}</SelectItem>
					) : null}
					{shown === STYLE_MIXED ? <SelectItem value={STYLE_MIXED}>{t("previewDocxStyleMixed")}</SelectItem> : null}
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
						closeOnClick
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
	const slotLabels = useSyncExternalStore(subscribeSlotLabels, slotLabelsShown)
	const locked = snapshot?.selection.locked ?? null
	const pictureItems = DRAWING_GROUP.length > 0 && snapshot?.context.drawing != null ? drawingItems(DRAWING_GROUP, props.surface) : []
	// Badges, the markup menu (the Review menu holds it once offered) and the slots take room from the groups.
	const fit = toolbarFit(
		width,
		(props.viewOnly.length > 0 ? 100 : 0) +
			(props.fileMenu === undefined || props.fileMenu === null ? 0 : 32) +
			(props.encrypted === undefined ? 0 : 140) +
			(locked === null ? 0 : 120) +
			(props.markup === null || REVIEW_SLOT.length > 0 ? 0 : 40) +
			slotsWidth(
				{
					insert: INSERT_SLOT.length > 0,
					review: REVIEW_SLOT.length > 0,
					table: TABLE_SLOT.length > 0 && snapshot?.context.table != null
				},
				slotLabels
			),
		FORMAT_SLOT.length > 0,
		FORMAT.length - KEPT_FORMAT_BUTTONS,
		pictureItems.length
	)
	const formatShown = KEPT_FORMAT_BUTTONS + fit.format
	const lineSpacing = single(snapshot?.values.lineSpacing ?? null)
	const percent = Math.round(props.zoom * 100)
	const commandProps = { snapshot, readOnly, execute }
	const commandOff = (type: DocxEditorCommandInfo["type"]): boolean => snapshot === undefined || readOnly || !snapshot.enabled[type]
	const slotProps = (commands: readonly DocxToolbarCommand[]): DocxToolbarSlotProps => ({
		// Mobile view shows no headers or footers: it does not offer editing them.
		commands: reflowCommands(commands, props.surface === "reflow"),
		snapshot,
		readOnly,
		apple,
		execute,
		surface: props.surface,
		inputRequest: props.inputRequest,
		requestInput: props.requestInput,
		onInputDone: props.onInputDone
	})
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

	if (props.viewOnly.length > 0 || props.encrypted !== undefined || locked !== null) {
		badges = (
			<div className="flex shrink-0 items-center gap-1">
				{props.viewOnly.length > 0 ? <ViewOnlyBadge reasons={props.viewOnly} /> : null}
				{props.encrypted === undefined ? null : (
					<PasswordBadge
						integrity={props.encrypted.integrity}
						saving={props.encrypted.saving}
					/>
				)}
				{locked === null ? null : <LockedBadge reason={locked} />}
			</div>
		)
	}

	return (
		<div
			ref={setRow}
			role="toolbar"
			aria-label={t("previewDocxEditToolbar")}
			className="flex h-10 shrink-0 items-center gap-0.5 border-b border-border px-2"
		>
			{props.fileMenu}
			{/* The groups, fitted to the row (toolbarFit); what still does not fit (badges, a Picture group on a
			phone) scrolls here, never pushing the ⋯ menu out of the row. */}
			<div className="-mx-1 flex min-w-0 flex-1 [scrollbar-width:none] items-center gap-0.5 self-stretch overflow-x-auto px-1 [&::-webkit-scrollbar]:hidden">
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
									styles={props.styles}
									style={snapshot?.paragraph.style ?? null}
									styleName={snapshot?.paragraph.styleName ?? ""}
									disabled={snapshot === undefined || readOnly || !snapshot.enabled["style.set"]}
									execute={execute}
								/>
							</>
						) : null}
						<Divider />
						{FORMAT.slice(0, formatShown).map(button)}
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
						{FORMAT_SLOT.length > 0 ? (
							<DocxFormatControls
								{...slotProps(FORMAT_SLOT)}
								fonts={props.fonts}
								theme={props.theme}
								painter={props.painter}
								onPainterCancel={props.onPainterCancel}
								showFont={fit.font}
								showColor={fit.color}
								showPainter={fit.painter}
							/>
						) : null}
						{INSERT_SLOT.length > 0 ? <DocxInsertMenu {...slotProps(INSERT_SLOT)} /> : null}
						{TABLE_SLOT.length > 0 && snapshot?.context.table != null ? (
							<DocxTableMenu
								{...slotProps(TABLE_SLOT)}
								theme={props.theme}
							/>
						) : null}
						{DRAWING_GROUP.length > 0 && snapshot?.context.drawing != null ? (
							<DocxDrawingMenu
								{...slotProps(DRAWING_GROUP)}
								items={pictureItems.slice(0, fit.picture)}
							/>
						) : null}
						<div className="min-w-0 flex-1" />
						{badges}
						{REVIEW_SLOT.length > 0 ? (
							<DocxReviewMenu
								{...slotProps(REVIEW_SLOT)}
								markup={props.markup}
								onMarkupChange={props.onMarkupChange}
							/>
						) : null}
						{props.markup === null || REVIEW_SLOT.length > 0 ? null : (
							<MarkupMenu
								mode={props.markup}
								onChange={props.onMarkupChange}
							/>
						)}
						{fit.find ? (
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
						) : null}
					</>
				)}
			</div>
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
					{!find.open && fit.picture < pictureItems.length ? (
						<DocxDrawingOverflow
							{...slotProps(DRAWING_GROUP)}
							items={pictureItems.slice(fit.picture)}
						/>
					) : null}
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
									{props.styles.map(option => {
										const label = styleLabel(option)

										return (
											<DropdownMenuRadioItem
												key={option}
												value={option}
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
					{!find.open && formatShown < FORMAT.length ? (
						<OverflowCommands
							commands={FORMAT.slice(formatShown)}
							{...commandProps}
						/>
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
					{!find.open && !(fit.style && fit.clear && fit.align && fit.lists && formatShown === FORMAT.length) ? (
						<DropdownMenuSeparator />
					) : null}
					{fit.find ? null : (
						<DropdownMenuItem
							onClick={() => {
								props.onFindOpen(false)
							}}
						>
							{t("previewDocxFindAction")}
						</DropdownMenuItem>
					)}
					<DropdownMenuItem
						onClick={() => {
							props.onFindOpen(true)
						}}
					>
						{t("previewDocxFindReplaceAction")}
					</DropdownMenuItem>
					{LINE_SPACING ? (
						<DropdownMenuSub>
							<DropdownMenuSubTrigger disabled={commandOff("paragraph.lineSpacing")}>
								{t("previewDocxLineSpacing")}
							</DropdownMenuSubTrigger>
							<DropdownMenuSubContent>
								<DropdownMenuRadioGroup
									value={lineSpacing?.rule === "auto" ? String(lineSpacing.value) : ""}
									onValueChange={(value: string) => {
										execute({ type: "paragraph.lineSpacing", rule: "auto", value: Number(value) })
									}}
								>
									{LINE_SPACINGS.map(value => (
										<DropdownMenuRadioItem
											key={value}
											value={String(value)}
											closeOnClick
										>
											{t("previewDocxLineSpacingValue", {
												value: value / 240,
												formatParams: { value: { minimumFractionDigits: 1 } }
											})}
										</DropdownMenuRadioItem>
									))}
								</DropdownMenuRadioGroup>
							</DropdownMenuSubContent>
						</DropdownMenuSub>
					) : null}
					{PARAGRAPH_DIALOG ? (
						<DropdownMenuItem
							disabled={commandOff("paragraph.indent") && commandOff("paragraph.spacing")}
							onClick={() => {
								props.requestInput("paragraph.indent")
							}}
						>
							{t("previewDocxParagraphAction")}
						</DropdownMenuItem>
					) : null}
					{PAGE_SETUP ? (
						<DropdownMenuItem
							disabled={commandOff("page.setup")}
							onClick={() => {
								props.requestInput("page.setup")
							}}
						>
							{t("previewDocxPageSetupAction")}
						</DropdownMenuItem>
					) : null}
					<DropdownMenuCheckboxItem
						closeOnClick
						checked={props.fieldCodes}
						disabled={snapshot?.enabled["view.fieldCodes"] !== true}
						onCheckedChange={() => {
							execute({ type: "view.fieldCodes" })
						}}
					>
						{t("previewDocxShowFieldCodes")}
					</DropdownMenuCheckboxItem>
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
					{VIEW_SLOT.length > 0 ? (
						<DocxReflowMenuItem
							{...slotProps(VIEW_SLOT)}
							reflow={props.reflow}
							onReflowChange={props.onReflowChange}
						/>
					) : null}
					<DropdownMenuSeparator />
					<DropdownMenuCheckboxItem
						closeOnClick
						checked={props.autosave}
						onCheckedChange={checked => {
							props.onAutosaveChange(checked)
						}}
					>
						{t("previewDocxAutosave")}
					</DropdownMenuCheckboxItem>
				</DropdownMenuContent>
			</DropdownMenu>
			{fit.zoom ? (
				<Tooltip>
					<TooltipTrigger
						render={
							<Button
								variant="ghost"
								size="xs"
								aria-label={t("previewDocxFitWidthAction", { percent })}
								className="w-12 shrink-0 px-0 text-muted-foreground tabular-nums"
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
			) : null}
		</div>
	)
}
