import { useEffect, useRef, useState, type KeyboardEvent } from "react"
import { useTranslation } from "react-i18next"
import {
	CaseSensitiveIcon,
	ChevronDownIcon,
	ChevronLeftIcon,
	ChevronRightIcon,
	ChevronUpIcon,
	EyeIcon,
	MessageSquareTextIcon,
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
import { viewOnlyReasonKey, type DocxViewOnlyReason } from "@/features/preview/components/docxViewer.logic"
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
import type { PreviewKey } from "@/lib/i18n"

// The DOCX viewer's toolbar row, in the PDF viewer's style: page navigation and zoom in the centre (the
// zoom percentage fits the page to the width again), swapped for an inline find field while finding; the
// view-only badge on the left, the markup menu and find on the right. Everything it shows comes from the
// viewer's command state; everything it does is a viewer command.

// Typing waits this long before the document is searched (each search runs over every page).
const FIND_DEBOUNCE_MS = 150

const MARKUP_LABELS: Readonly<Record<DocxMarkupMode, PreviewKey>> = {
	all: "previewDocxMarkupAll",
	simple: "previewDocxMarkupSimple",
	none: "previewDocxMarkupNone",
	original: "previewDocxMarkupOriginal"
}

export interface DocxToolbarProps {
	state: DocxViewerState
	dispatch: (command: DocxViewerCommand) => void
	// Null for a document without comments or tracked changes.
	markup: DocxMarkupMode | null
	onMarkupChange: (mode: DocxMarkupMode) => void
	viewOnly: readonly DocxViewOnlyReason[]
	// Bumped to focus the find field (and select its text) again while it is open.
	findFocus: number
	// Where focus goes when the find field closes.
	onFindClosed: () => void
}

function ViewOnlyBadge({ reasons }: { reasons: readonly DocxViewOnlyReason[] }) {
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

function MarkupMenu({ mode, onChange }: { mode: DocxMarkupMode; onChange: (mode: DocxMarkupMode) => void }) {
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

function FindField({
	state,
	dispatch,
	focus,
	onClosed
}: {
	state: DocxViewerState
	dispatch: (command: DocxViewerCommand) => void
	focus: number
	onClosed: () => void
}) {
	const { t } = useTranslation("preview")
	const find = state.find
	const [text, setText] = useState(find.query)
	const timer = useRef<ReturnType<typeof setTimeout> | null>(null)
	const inputRef = useRef<HTMLInputElement | null>(null)
	const searched = text === find.query
	const current = find.count === 0 ? 0 : find.active + 1

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

		dispatch({ type: "find.query", query, caseSensitive })
	}

	const close = (): void => {
		dispatch({ type: "find.close" })
		onClosed()
	}

	const onKeyDown = (event: KeyboardEvent<HTMLInputElement>): void => {
		// Keys an input method is composing with are its own (Enter confirms the composition).
		if (isImeKeydown(event.nativeEvent)) {
			return
		}

		if (event.key === "Escape") {
			// Taken here: the preview overlay closes on an Escape nothing inside took.
			event.preventDefault()
			close()
		} else if (event.key === "Enter") {
			event.preventDefault()

			if (!searched) {
				search(text, find.caseSensitive)
			} else {
				dispatch({ type: event.shiftKey ? "find.previous" : "find.next" })
			}
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
					className="h-7 w-40 min-w-0 rounded-xl bg-muted/60 pr-2 pl-7 text-xs leading-5 outline-none placeholder:text-muted-foreground focus-visible:ring-[3px] focus-visible:ring-ring/30 max-sm:w-28 [&::-webkit-search-cancel-button]:appearance-none"
					onChange={event => {
						const query = event.target.value

						setText(query)

						if (timer.current !== null) {
							clearTimeout(timer.current)
						}

						timer.current = setTimeout(() => {
							timer.current = null
							dispatch({ type: "find.query", query, caseSensitive: find.caseSensitive })
						}, FIND_DEBOUNCE_MS)
					}}
					onKeyDown={onKeyDown}
				/>
			</div>
			<span
				aria-hidden="true"
				className="w-12 shrink-0 text-center text-xs text-muted-foreground tabular-nums"
			>
				{text.length > 0 && searched ? t("previewDocxFindCount", { current, total: find.count }) : null}
			</span>
			<span
				role="status"
				className="sr-only"
			>
				{text.length > 0 && searched
					? find.count === 0
						? t("previewDocxFindNoMatches")
						: t("previewDocxFindStatus", { current, total: find.count })
					: null}
			</span>
			<TooltipIconButton
				label={t("previewDocxFindPrevious")}
				disabled={!isDocxCommandEnabled(state, "find.previous")}
				onClick={() => {
					dispatch({ type: "find.previous" })
				}}
			>
				<ChevronUpIcon />
			</TooltipIconButton>
			<TooltipIconButton
				label={t("previewDocxFindNext")}
				disabled={!isDocxCommandEnabled(state, "find.next")}
				onClick={() => {
					dispatch({ type: "find.next" })
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

export function DocxToolbar({ state, dispatch, markup, onMarkupChange, viewOnly, findFocus, onFindClosed }: DocxToolbarProps) {
	const { t } = useTranslation("preview")

	return (
		<div className="grid h-10 shrink-0 grid-cols-[1fr_auto_1fr] items-center gap-2 px-2">
			<div className="flex min-w-0 items-center">{viewOnly.length > 0 ? <ViewOnlyBadge reasons={viewOnly} /> : null}</div>
			{state.find.open ? (
				<FindField
					state={state}
					dispatch={dispatch}
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
