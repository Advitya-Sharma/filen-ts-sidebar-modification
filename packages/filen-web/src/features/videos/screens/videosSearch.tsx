import { useEffect, useId, useMemo, useState, type KeyboardEvent } from "react"
import { useTranslation } from "react-i18next"
import { useRouter } from "@tanstack/react-router"
import { SearchIcon, SearchXIcon } from "lucide-react"
import { useElementSize } from "@/lib/useElementSize"
import { SearchInput } from "@/features/drive/components/searchInput"
import { columnsForWidth } from "@/features/drive/lib/gridLayout"
import { useVideoHomeData } from "@/features/videos/hooks/useVideoHomeData"
import { useVideoCardActions } from "@/features/videos/hooks/useVideoCardActions"
import { useCardFocus } from "@/features/videos/hooks/useCardFocus"
import { CardObserver } from "@/features/videos/lib/cardObserver"
import { itemCard, titleCard, type HomeData, type VideoCardModel } from "@/features/videos/lib/homeRows.logic"
import { SEARCH_DEBOUNCE_MS, buildVideoSearchIndex, searchVideos, type SearchResultGroup } from "@/features/videos/lib/search.logic"
import { gridKeyAction } from "@/features/videos/lib/videoGrid.logic"
import { VideoCard, type VideoCardAction } from "@/features/videos/components/videoCard"
import { EmptyMessage } from "@/components/emptyMessage"

const GAP = 16
const POSTER_MIN = 132
const LANDSCAPE_MIN = 220

const GROUP_LABEL_KEYS = {
	movies: "videosMovies",
	shows: "videosShows",
	episodes: "videosSearchEpisodes",
	other: "videosOther"
} as const

function groupCards(group: SearchResultGroup, data: Pick<HomeData, "index" | "entries" | "watch">): VideoCardModel[] {
	return group.results.map(result =>
		result.item === null
			? titleCard({ data, title: result.title, shape: result.title.kind === "other" ? "landscape" : "poster" })
			: itemCard({ data, item: result.item, title: result.title })
	)
}

interface ResultGroupProps {
	group: SearchResultGroup
	cards: VideoCardModel[]
	width: number
	observer: CardObserver | null
	onAction: (action: VideoCardAction, card: VideoCardModel) => void
}

// One group of results: a labelled grid with its own roving tab stop.
function ResultGroup({ group, cards, width, observer, onAction }: ResultGroupProps) {
	const { t } = useTranslation("videos")
	const headingId = useId()
	const router = useRouter()
	const [active, setActive] = useState(0)
	const cardFocus = useCardFocus<number>()
	const landscape = group.kind === "episodes" || group.kind === "other"
	const columns = columnsForWidth(width, landscape ? LANDSCAPE_MIN : POSTER_MIN, GAP)
	const activeIndex = Math.min(active, Math.max(0, cards.length - 1))

	function handleKeyDown(event: KeyboardEvent<HTMLDivElement>): void {
		const target = event.target

		if (
			!(target instanceof HTMLAnchorElement) ||
			!target.hasAttribute("data-video-card") ||
			event.altKey ||
			event.ctrlKey ||
			event.metaKey
		) {
			return
		}

		const action = gridKeyAction({ key: event.key, index: activeIndex, count: cards.length, columns, pageRows: 2 })

		if (action.kind === "none" || action.kind === "open") {
			return
		}

		if (action.kind === "back") {
			if (router.history.canGoBack()) {
				event.preventDefault()
				router.history.back()
			}

			return
		}

		event.preventDefault()

		if (action.kind === "play") {
			const card = cards[activeIndex]

			if (card !== undefined) {
				onAction("play", card)
			}

			return
		}

		setActive(action.index)
		cardFocus.focus(action.index, { scroll: true })
	}

	return (
		<section
			aria-labelledby={headingId}
			className="flex flex-col gap-2"
		>
			<div className="flex items-baseline justify-between gap-3">
				<h2
					id={headingId}
					className="text-sm font-medium"
				>
					{t(GROUP_LABEL_KEYS[group.kind])}
				</h2>
				{group.total > group.results.length ? (
					<span className="text-xs text-muted-foreground tabular-nums">
						{t("videosSearchShownOf", { shown: group.results.length, total: group.total })}
					</span>
				) : null}
			</div>
			<div
				role="list"
				aria-labelledby={headingId}
				className="grid py-2"
				style={{ gridTemplateColumns: `repeat(${String(columns)}, minmax(0, 1fr))`, gap: GAP }}
				onKeyDown={handleKeyDown}
			>
				{cards.map((card, index) => (
					<VideoCard
						key={card.key}
						card={card}
						nearObserver={observer}
						active={index === activeIndex}
						registerRef={element => {
							cardFocus.register(index, element)
						}}
						onFocusCard={() => {
							setActive(index)
						}}
						onAction={onAction}
					/>
				))}
			</div>
		</section>
	)
}

// Videos search: a debounced query over an in-memory index of what the library holds, grouped Movies, Shows,
// Episodes and Other. Nothing is listed or probed to answer it; card art loads as the cards come near.
export function VideosSearchScreen() {
	const { t } = useTranslation("videos")
	const data = useVideoHomeData()
	const onAction = useVideoCardActions(data.items)
	const [input, setInput] = useState("")
	const [query, setQuery] = useState("")
	const [scrollElement, setScrollElement] = useState<HTMLDivElement | null>(null)
	const [toolbar, setToolbar] = useState<HTMLDivElement | null>(null)
	const { width } = useElementSize(scrollElement)

	useEffect(() => {
		toolbar?.querySelector("input")?.focus()
	}, [toolbar])

	useEffect(() => {
		const timer = setTimeout(() => {
			setQuery(input)
		}, SEARCH_DEBOUNCE_MS)

		return () => {
			clearTimeout(timer)
		}
	}, [input])

	// Cards within half a screen of the visible part.
	const observer = useMemo(
		() => (scrollElement === null ? null : new CardObserver({ root: scrollElement, rootMargin: "50% 0px", threshold: 0 })),
		[scrollElement]
	)

	useEffect(() => {
		return () => {
			observer?.disconnect()
		}
	}, [observer])

	// No hook call below: the compiler keeps these derivations memoized only while none follows them.
	const index = buildVideoSearchIndex({ titles: data.snapshot.titles, items: data.items ?? [] })
	const groups = searchVideos({ index, query })
	const homeData: Pick<HomeData, "index" | "entries" | "watch"> = { index: data.index, entries: data.entries, watch: data.watch }
	const cards = groups.map(group => groupCards(group, homeData))
	const total = groups.reduce((sum, group) => sum + group.total, 0)
	const trimmed = query.trim()

	return (
		<div className="flex min-h-0 flex-1 flex-col overflow-hidden">
			<div
				ref={setToolbar}
				className="flex shrink-0 items-center gap-3 px-4 pb-3"
			>
				<SearchInput
					action="videos.search"
					label={t("videosSearchPlaceholder")}
					value={input}
					onChange={setInput}
					onClear={() => {
						setInput("")
					}}
				/>
				<span
					role="status"
					className="ml-auto text-xs text-muted-foreground tabular-nums"
				>
					{trimmed.length > 0 ? t("videosSearchResults", { count: total }) : ""}
				</span>
			</div>
			<div
				ref={setScrollElement}
				className="flex min-h-0 flex-1 flex-col gap-6 overflow-y-auto px-4 pb-6"
			>
				{trimmed.length === 0 ? (
					<EmptyMessage
						icon={SearchIcon}
						title={t("videosSearchEmptyTitle")}
						description={t("videosSearchEmptyBody")}
					/>
				) : groups.length === 0 ? (
					<EmptyMessage
						icon={SearchXIcon}
						title={t("videosSearchNoResultsTitle", { query: trimmed })}
						description={t("videosSearchNoResultsBody")}
					/>
				) : width > 0 ? (
					groups.map((group, position) => (
						<ResultGroup
							key={group.kind}
							group={group}
							cards={cards[position] ?? []}
							width={width}
							observer={observer}
							onAction={onAction}
						/>
					))
				) : null}
			</div>
		</div>
	)
}
