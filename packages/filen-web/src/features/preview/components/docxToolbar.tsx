import { useEffect, useId, useRef, useState, type KeyboardEvent } from "react"
import { useTranslation } from "react-i18next"
import {
	CaseSensitiveIcon,
	ChevronDownIcon,
	ChevronLeftIcon,
	ChevronRightIcon,
	ChevronUpIcon,
	EyeIcon,
	LockIcon,
	TriangleAlertIcon,
	MessageSquareTextIcon,
	ReplaceAllIcon,
	ReplaceIcon,
	SearchIcon,
	XIcon,
	ZoomInIcon,
	ZoomOutIcon
} from "lucide-react"
import {
	DOCX_MARKUP_MODES,
	isDocxCommandEnabled,
	type DocxMarkupMode,
	type DocxViewerCommand,
	type DocxViewerState
} from "@filen/office/dom/docx"
import { cn } from "@filen/shared"
import {
	MARKUP_LABELS,
	viewOnlyReasonKey,
	viewerFindControls,
	type DocxFindControls,
	type DocxViewOnlyReason
} from "@/features/preview/components/docxViewer.logic"
import { Badge } from "@/components/ui/badge"
import { Button } from "@/components/ui/button"
import {
	DropdownMenu,
	DropdownMenuContent,
	DropdownMenuLabel,
	DropdownMenuRadioGroup,
	DropdownMenuRadioItem,
	DropdownMenuTrigger
} from "@/components/ui/dropdown-menu"
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover"
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip"
import { TooltipIconButton } from "@/components/ui/tooltipIconButton"
import { Kbd } from "@/lib/keymap/kbd"
import { useAction } from "@/lib/keymap/useAction"
import { isImeKeydown } from "@/lib/ime"

// The DOCX viewer's toolbar row, in the PDF viewer's style: page navigation and zoom in the centre (the
// zoom percentage fits the page to the width again), swapped for an inline find field while finding; the
// view-only badge on the left, the markup menu and find on the right. Everything it shows comes from the
// viewer's command state; everything it does is a viewer command.

// Typing waits this long before the document is searched (each search runs over every page).
const FIND_DEBOUNCE_MS = 150
const FIELD_CLASS =
	"h-7 w-40 min-w-0 rounded-xl bg-muted/60 pr-2 pl-7 text-xs leading-5 outline-none placeholder:text-muted-foreground focus-visible:ring-[3px] focus-visible:ring-ring/30 disabled:opacity-50 max-sm:w-28 [&::-webkit-search-cancel-button]:appearance-none"

export interface DocxToolbarProps {
	state: DocxViewerState
	dispatch: (command: DocxViewerCommand) => void
	// Null for a document without comments or tracked changes.
	markup: DocxMarkupMode | null
	onMarkupChange: (mode: DocxMarkupMode) => void
	viewOnly: readonly DocxViewOnlyReason[]
	// Opened with its password, and whether its integrity was verified.
	encrypted?: PasswordIntegrity | undefined
	// Bumped to focus the find field (and select its text) again while it is open.
	findFocus: number
	// Where focus goes when the find field closes.
	onFindClosed: () => void
}

export function ViewOnlyBadge({ reasons }: { reasons: readonly DocxViewOnlyReason[] }) {
	const { t } = useTranslation("preview")

	return (
		<Popover>
			<Badge
				variant="outline"
				className="cursor-pointer text-muted-foreground hover:bg-muted"
				render={<PopoverTrigger aria-label={t("previewDocxViewOnly")} />}
			>
				<EyeIcon data-icon="inline-start" />
				<span className="max-sm:hidden">{t("previewDocxViewOnly")}</span>
			</Badge>
			<PopoverContent
				align="start"
				className="gap-2"
			>
				<p className="font-medium">{t("previewDocxViewOnlyTitle")}</p>
				<ul className="flex list-disc flex-col gap-1 pl-4 text-muted-foreground">
					{reasons.map(reason => (
						<li key={reason}>{t(viewOnlyReasonKey(reason))}</li>
					))}
				</ul>
			</PopoverContent>
		</Popover>
	)
}

// Whether an encrypted package's integrity was verified ([MS-OFFCRYPTO] §2.3.4.14 dataIntegrity HMAC), or could
// not be: Standard encryption has none, and an Agile file may leave it out, so its content may have been changed.
export type PasswordIntegrity = "verified" | "unchecked"

// A document opened with its password: `saving` while its saves keep the password (an editor that can save),
// else it was decrypted here only. An unchecked one says so, in the popover and as the badge's description.
export function PasswordBadge({ integrity, saving }: { integrity: PasswordIntegrity; saving: boolean }) {
	const { t } = useTranslation("preview")
	const noteId = useId()
	const unchecked = integrity === "unchecked"

	return (
		<Popover>
			<Badge
				variant="outline"
				className={cn("cursor-pointer hover:bg-muted", unchecked ? "text-warning-foreground" : "text-muted-foreground")}
				render={
					<PopoverTrigger
						aria-label={t("previewDocxPasswordProtected")}
						aria-describedby={unchecked ? noteId : undefined}
					/>
				}
			>
				{unchecked ? <TriangleAlertIcon data-icon="inline-start" /> : <LockIcon data-icon="inline-start" />}
				<span className="max-sm:hidden">{t("previewDocxPasswordProtected")}</span>
			</Badge>
			{unchecked ? (
				<span
					id={noteId}
					className="sr-only"
				>
					{t("previewDocxPasswordUnchecked")}
				</span>
			) : null}
			<PopoverContent
				align="start"
				className="gap-2"
			>
				<p className="font-medium">{t("previewDocxPasswordProtected")}</p>
				<p className="text-muted-foreground">{t(saving ? "previewDocxPasswordSaved" : "previewDocxPasswordOpened")}</p>
				{unchecked ? <p className="text-warning-foreground">{t("previewDocxPasswordUnchecked")}</p> : null}
			</PopoverContent>
		</Popover>
	)
}

export function MarkupMenu({
	mode,
	onChange,
	disabled
}: {
	mode: DocxMarkupMode
	onChange: (mode: DocxMarkupMode) => void
	// Modes that cannot be switched to now (another layout while edits are unsaved).
	disabled?: (mode: DocxMarkupMode) => boolean
}) {
	const { t } = useTranslation("preview")

	return (
		<DropdownMenu>
			<Tooltip>
				<TooltipTrigger
					render={
						<DropdownMenuTrigger
							render={
								<Button
									variant="ghost"
									size="sm"
									aria-label={`${t("previewDocxMarkup")}: ${t(MARKUP_LABELS[mode])}`}
									className="px-2"
								>
									<MessageSquareTextIcon />
									<span className="text-xs max-md:hidden">{t(MARKUP_LABELS[mode])}</span>
								</Button>
							}
						/>
					}
				/>
				<TooltipContent>{t("previewDocxMarkup")}</TooltipContent>
			</Tooltip>
			<DropdownMenuContent
				align="end"
				className="w-auto min-w-40"
			>
				<DropdownMenuRadioGroup
					value={mode}
					onValueChange={(next: DocxMarkupMode) => {
						onChange(next)
					}}
				>
					<DropdownMenuLabel>{t("previewDocxMarkup")}</DropdownMenuLabel>
					{DOCX_MARKUP_MODES.map(info => (
						<DropdownMenuRadioItem
							key={info.mode}
							value={info.mode}
							disabled={disabled?.(info.mode) === true}
							closeOnClick
						>
							{t(MARKUP_LABELS[info.mode])}
						</DropdownMenuRadioItem>
					))}
				</DropdownMenuRadioGroup>
			</DropdownMenuContent>
		</DropdownMenu>
	)
}

export function DocxFindField({
	find,
	focus,
	onClosed,
	replaceOpen: replaceOpenProp
}: {
	find: DocxFindControls
	// Bumped to focus the find field (and select its text) again while it is open.
	focus: number
	onClosed: () => void
	// Opened as "Find and replace": the replace field shows from the start.
	replaceOpen?: boolean
}) {
	const { t } = useTranslation("preview")
	const [text, setText] = useState(find.query)
	const [replacement, setReplacement] = useState("")
	const [replaceOpen, setReplaceOpen] = useState(replaceOpenProp === true)
	const timer = useRef<ReturnType<typeof setTimeout> | null>(null)
	const inputRef = useRef<HTMLInputElement | null>(null)
	const searched = text === find.query
	const current = find.count === 0 ? 0 : find.active + 1
	const counts = { current, total: find.count }
	const replace = find.replace

	// Mod+F again in the field selects what it holds, as in a browser's find bar.
	const findAgain = useAction(
		"editor.find",
		keyboardEvent => {
			keyboardEvent.preventDefault()
			inputRef.current?.select()
		},
		{ enableOnFormTags: ["input"] }
	)

	useEffect(() => {
		inputRef.current?.focus()
		inputRef.current?.select()
	}, [focus])

	useEffect(() => {
		return () => {
			if (timer.current !== null) {
				clearTimeout(timer.current)
			}
		}
	}, [])

	const search = (query: string, caseSensitive: boolean): void => {
		if (timer.current !== null) {
			clearTimeout(timer.current)
			timer.current = null
		}

		find.search(query, caseSensitive)
	}

	const close = (): void => {
		find.close()
		onClosed()
	}

	// Keys an input method is composing with are its own (Enter confirms the composition); Escape is taken
	// here, as the preview overlay closes on an Escape nothing inside took.
	const onFieldKeyDown = (event: KeyboardEvent<HTMLInputElement>, enter: (shift: boolean) => void): void => {
		if (isImeKeydown(event.nativeEvent)) {
			return
		}

		if (event.key === "Escape") {
			event.preventDefault()
			close()
		} else if (event.key === "Enter") {
			event.preventDefault()
			enter(event.shiftKey)
		}
	}

	return (
		<div className="flex min-w-0 items-center gap-1">
			<div className="relative min-w-0">
				<SearchIcon
					aria-hidden="true"
					className="pointer-events-none absolute top-1/2 left-2 size-3.5 -translate-y-1/2 text-muted-foreground"
				/>
				<input
					ref={element => {
						inputRef.current = element
						findAgain(element)
					}}
					type="search"
					aria-label={t("previewDocxFindPlaceholder")}
					placeholder={t("previewDocxFindPlaceholder")}
					value={text}
					className={FIELD_CLASS}
					onChange={event => {
						const query = event.target.value

						setText(query)

						if (timer.current !== null) {
							clearTimeout(timer.current)
						}

						timer.current = setTimeout(() => {
							timer.current = null
							find.search(query, find.caseSensitive)
						}, FIND_DEBOUNCE_MS)
					}}
					onKeyDown={event => {
						onFieldKeyDown(event, shift => {
							if (!searched) {
								search(text, find.caseSensitive)
							} else {
								find.step(shift ? -1 : 1)
							}
						})
					}}
				/>
			</div>
			<span
				aria-hidden="true"
				className="min-w-12 shrink-0 text-center text-xs text-muted-foreground tabular-nums"
			>
				{text.length > 0 && searched ? t(find.capped ? "previewDocxFindCountCapped" : "previewDocxFindCount", counts) : null}
			</span>
			<span
				role="status"
				className="sr-only"
			>
				{text.length > 0 && searched
					? find.count === 0
						? t("previewDocxFindNoMatches")
						: t(find.capped ? "previewDocxFindStatusCapped" : "previewDocxFindStatus", counts)
					: null}
			</span>
			<TooltipIconButton
				label={t("previewDocxFindPrevious")}
				disabled={find.count === 0}
				onClick={() => {
					find.step(-1)
				}}
			>
				<ChevronUpIcon />
			</TooltipIconButton>
			<TooltipIconButton
				label={t("previewDocxFindNext")}
				disabled={find.count === 0}
				onClick={() => {
					find.step(1)
				}}
			>
				<ChevronDownIcon />
			</TooltipIconButton>
			<TooltipIconButton
				label={t("previewDocxFindMatchCase")}
				pressed={find.caseSensitive}
				className={find.caseSensitive ? "bg-accent text-foreground" : undefined}
				onClick={() => {
					search(text, !find.caseSensitive)
				}}
			>
				<CaseSensitiveIcon />
			</TooltipIconButton>
			{replace === undefined ? null : (
				<>
					<TooltipIconButton
						label={t("previewDocxShowReplace")}
						pressed={replaceOpen}
						className={replaceOpen ? "bg-accent text-foreground" : undefined}
						onClick={() => {
							setReplaceOpen(open => !open)
						}}
					>
						<ReplaceIcon />
					</TooltipIconButton>
					{replaceOpen ? (
						<>
							<input
								type="text"
								aria-label={t("previewDocxReplacePlaceholder")}
								placeholder={t("previewDocxReplacePlaceholder")}
								value={replacement}
								disabled={replace.disabled}
								className={cn(FIELD_CLASS, "pl-2")}
								onChange={event => {
									setReplacement(event.target.value)
								}}
								onKeyDown={event => {
									onFieldKeyDown(event, () => {
										if (find.count > 0 && searched) {
											replace.one(replacement)
										}
									})
								}}
							/>
							<TooltipIconButton
								label={t("previewDocxReplace")}
								disabled={replace.disabled || find.count === 0 || !searched}
								onClick={() => {
									replace.one(replacement)
								}}
							>
								<ReplaceIcon />
							</TooltipIconButton>
							<TooltipIconButton
								label={t("previewDocxReplaceAll")}
								disabled={replace.disabled || find.count === 0 || !searched}
								onClick={() => {
									replace.all(replacement)
								}}
							>
								<ReplaceAllIcon />
							</TooltipIconButton>
						</>
					) : null}
				</>
			)}
			<TooltipIconButton
				label={t("previewDocxFindClose")}
				onClick={close}
			>
				<XIcon />
			</TooltipIconButton>
		</div>
	)
}

function Navigation({ state, dispatch }: { state: DocxViewerState; dispatch: (command: DocxViewerCommand) => void }) {
	const { t } = useTranslation("preview")
	const percent = Math.round(state.zoom * 100)

	return (
		<div className="flex min-w-0 items-center gap-2 max-sm:gap-1">
			<Button
				variant="ghost"
				size="icon-sm"
				disabled={!isDocxCommandEnabled(state, "page.previous")}
				aria-label={t("previewDocxPreviousPageAction")}
				onClick={() => {
					dispatch({ type: "page.previous" })
				}}
			>
				<ChevronLeftIcon />
			</Button>
			<span className="text-xs whitespace-nowrap text-muted-foreground tabular-nums">
				{t("previewDocxPageIndicator", { current: Math.min(state.page + 1, state.pageCount), total: state.pageCount })}
			</span>
			<Button
				variant="ghost"
				size="icon-sm"
				disabled={!isDocxCommandEnabled(state, "page.next")}
				aria-label={t("previewDocxNextPageAction")}
				onClick={() => {
					dispatch({ type: "page.next" })
				}}
			>
				<ChevronRightIcon />
			</Button>
			<span
				aria-hidden="true"
				className="mx-1 h-4 w-px bg-border max-sm:mx-0"
			/>
			<Button
				variant="ghost"
				size="icon-sm"
				disabled={!isDocxCommandEnabled(state, "zoom.out")}
				aria-label={t("previewDocxZoomOutAction")}
				onClick={() => {
					dispatch({ type: "zoom.out" })
				}}
			>
				<ZoomOutIcon />
			</Button>
			<Tooltip>
				<TooltipTrigger
					render={
						<Button
							variant="ghost"
							size="xs"
							aria-label={t("previewDocxFitWidthAction", { percent })}
							className="w-12 px-0 text-muted-foreground tabular-nums"
							onClick={() => {
								dispatch({ type: "zoom.fitWidth" })
							}}
						>
							{t("previewDocxZoomIndicator", { percent })}
						</Button>
					}
				/>
				<TooltipContent>{t("previewDocxFitWidth")}</TooltipContent>
			</Tooltip>
			<Button
				variant="ghost"
				size="icon-sm"
				disabled={!isDocxCommandEnabled(state, "zoom.in")}
				aria-label={t("previewDocxZoomInAction")}
				onClick={() => {
					dispatch({ type: "zoom.in" })
				}}
			>
				<ZoomInIcon />
			</Button>
		</div>
	)
}

export function DocxToolbar({ state, dispatch, markup, onMarkupChange, viewOnly, encrypted, findFocus, onFindClosed }: DocxToolbarProps) {
	const { t } = useTranslation("preview")

	return (
		<div className="grid h-10 shrink-0 grid-cols-[1fr_auto_1fr] items-center gap-2 px-2">
			<div className="flex min-w-0 items-center gap-1">
				{viewOnly.length > 0 ? <ViewOnlyBadge reasons={viewOnly} /> : null}
				{encrypted === undefined ? null : (
					<PasswordBadge
						integrity={encrypted}
						saving={false}
					/>
				)}
			</div>
			{state.find.open ? (
				<DocxFindField
					find={viewerFindControls(state, dispatch)}
					focus={findFocus}
					onClosed={onFindClosed}
				/>
			) : (
				<Navigation
					state={state}
					dispatch={dispatch}
				/>
			)}
			<div className="flex min-w-0 items-center justify-end gap-1">
				{markup === null ? null : (
					<MarkupMenu
						mode={markup}
						onChange={onMarkupChange}
					/>
				)}
				{state.find.open ? null : (
					<TooltipIconButton
						label={t("previewDocxFindAction")}
						disabled={!isDocxCommandEnabled(state, "find.open")}
						shortcut={<Kbd action="editor.find" />}
						onClick={() => {
							dispatch({ type: "find.open" })
						}}
					>
						<SearchIcon />
					</TooltipIconButton>
				)}
			</div>
		</div>
	)
}
