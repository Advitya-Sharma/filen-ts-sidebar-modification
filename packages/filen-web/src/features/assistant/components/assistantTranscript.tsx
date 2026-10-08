import { useCallback, useRef } from "react"
import { useTranslation } from "react-i18next"
import { useVirtualizer, type VirtualizerOptions } from "@tanstack/react-virtual"
import { cn } from "@filen/shared"
import { AssistantReplyMarkdown } from "@/features/assistant/components/replyMarkdown"
import { BubbleTail, OTHER_TAIL_CLASS, OWN_TAIL_CLASS } from "@/features/chats/components/thread/messageRow"
import { useIsSidebarPanelVisible } from "@/features/shell/lib/sidebarPanelVisibility"
import { observeElementOffsetFromAttach } from "@/lib/virtualScroll"
import { AssistantConfirmationCard } from "@/features/assistant/components/assistantConfirmationCard"
import { AssistantSteps } from "@/features/assistant/components/assistantSteps"
import { TURN_END_KEYS, type TranscriptRow } from "@/features/assistant/components/assistant.logic"
import { stepsExpanded } from "@/features/assistant/store/transcript"

function estimateRowSize(row: TranscriptRow | undefined): number {
	switch (row?.kind) {
		case "reply":
			return 96
		case "proposal":
			return 200
		case "steps":
			return row.outcome === null || stepsExpanded(row) ? 12 + 22 * row.steps.length : 30
		case "user":
		case "thinking":
			return 48
		default:
			return 28
	}
}

// Shared with the chat thread's bubbles: same radius, tail and colour tokens, so the panel reads as a
// conversation the moment it opens.
const BUBBLE_CLASS = "relative min-w-0 rounded-[18px] px-3 py-1.5 text-sm"

// Rendered by a component the React Compiler does compile (AssistantTranscript opts out because of the
// virtualizer), so a scroll step or a token landing in the last reply reuses every other row's subtree.
function TranscriptRowContent({ row }: { row: TranscriptRow }) {
	const { t } = useTranslation("assistant")

	switch (row.kind) {
		case "user":
			return (
				<div className="flex justify-end px-4 py-1">
					<div
						className={cn(
							BUBBLE_CLASS,
							"max-w-[85%] rounded-br-[6px] bg-chat-own wrap-anywhere whitespace-pre-wrap text-chat-own-foreground select-text"
						)}
					>
						<span className="sr-only">{t("assistantYouSaid")}: </span>
						{row.text}
						<BubbleTail className={cn(OWN_TAIL_CLASS, "text-chat-own")} />
					</div>
				</div>
			)
		case "reply":
			return (
				<div className="flex justify-start px-4 py-1">
					<div
						aria-busy={!row.final}
						className={cn(BUBBLE_CLASS, "max-w-[92%] rounded-bl-[6px] bg-chat-other text-foreground")}
					>
						<AssistantReplyMarkdown
							text={row.withheld === true ? t("assistantReplyWithheld") : row.text}
							alt={t("assistantReplyLabel")}
							className="size-auto overflow-visible p-0 wrap-anywhere [&_p]:mb-2 [&_pre]:bg-background/60 [&>:last-child]:mb-0"
						/>
						<BubbleTail className={cn(OTHER_TAIL_CLASS, "text-chat-other")} />
					</div>
				</div>
			)
		case "steps":
			return <AssistantSteps entry={row} />
		case "proposal":
			return (
				<div className="px-4 py-1.5">
					<AssistantConfirmationCard
						proposal={row.proposal}
						status={row.status}
					/>
				</div>
			)
		case "turnEnd":
			return (
				<p
					className={cn(
						"px-5 py-0.5 text-xs leading-5",
						row.outcome === "cancelled" ? "text-muted-foreground" : "text-destructive"
					)}
				>
					{t(TURN_END_KEYS[row.outcome])}
				</p>
			)
		case "thinking":
			return (
				<div
					aria-hidden="true"
					className="flex px-4 py-1"
				>
					<div className="relative flex h-8 items-center gap-1 rounded-[18px] rounded-bl-[6px] bg-chat-other px-3.5">
						<span className="size-1.5 rounded-full bg-muted-foreground motion-safe:animate-bounce motion-safe:[animation-delay:-0.3s]" />
						<span className="size-1.5 rounded-full bg-muted-foreground motion-safe:animate-bounce motion-safe:[animation-delay:-0.15s]" />
						<span className="size-1.5 rounded-full bg-muted-foreground motion-safe:animate-bounce" />
						<BubbleTail className={`${OTHER_TAIL_CLASS} text-chat-other`} />
					</div>
				</div>
			)
	}
}

type TranscriptVirtualizerOptions = VirtualizerOptions<HTMLDivElement, HTMLDivElement>

const observeOffset: TranscriptVirtualizerOptions["observeElementOffset"] = observeElementOffsetFromAttach

// Virtualized like the chat thread, but top-origin: the virtualizer's own end anchoring keeps a reader who
// sits at the bottom pinned there as rows append and as the streaming reply grows, and leaves a reader
// who scrolled up where they are. Not rendered off screen: a hidden list would measure every row at zero.
export function AssistantTranscript({ rows }: { rows: readonly TranscriptRow[] }) {
	// Kept out of the React Compiler: the virtualizer is one mutable instance, so row JSX memoized on it
	// would go stale on scroll. Its two callbacks are memoized by hand instead.
	"use no memo"

	const { t } = useTranslation("assistant")
	const visible = useIsSidebarPanelVisible()
	const scrollRef = useRef<HTMLDivElement | null>(null)

	// The virtualizer re-lays every row when its key function changes, so it changes with the rows only.
	const getItemKey = useCallback((index: number) => rows[index]?.id ?? index, [rows])

	// Opens at the newest message: set before the virtualizer attaches, which then reads this offset.
	const attachScroller = useCallback((element: HTMLDivElement | null) => {
		scrollRef.current = element

		if (element !== null) {
			element.scrollTop = element.scrollHeight
		}
	}, [])

	const virtualizer = useVirtualizer<HTMLDivElement, HTMLDivElement>({
		count: rows.length,
		getScrollElement: () => scrollRef.current,
		estimateSize: index => estimateRowSize(rows[index]),
		getItemKey,
		overscan: 6,
		enabled: visible,
		observeElementOffset: observeOffset,
		anchorTo: "end",
		followOnAppend: true,
		scrollEndThreshold: 24,
		paddingStart: 8,
		paddingEnd: 8
	})

	// Focusable, so the keyboard can scroll it once it is clicked or tabbed into.
	return (
		<div
			ref={attachScroller}
			role="region"
			aria-label={t("assistantConversationLabel")}
			tabIndex={0}
			className="min-h-0 flex-1 overflow-y-auto focus-ring-row outline-none [overflow-anchor:none] focus-visible:ring-inset"
		>
			<div
				className="relative w-full"
				style={{ height: virtualizer.getTotalSize() }}
			>
				{virtualizer.getVirtualItems().map(virtualRow => {
					const row = rows[virtualRow.index]

					if (row === undefined) {
						return null
					}

					return (
						<div
							key={virtualRow.key}
							data-index={virtualRow.index}
							ref={virtualizer.measureElement}
							className="absolute top-0 left-0 w-full"
							style={{ transform: `translateY(${String(virtualRow.start)}px)` }}
						>
							<TranscriptRowContent row={row} />
						</div>
					)
				})}
			</div>
		</div>
	)
}
