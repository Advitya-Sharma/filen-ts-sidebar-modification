import {
	useCallback,
	useEffect,
	useLayoutEffect,
	useMemo,
	useRef,
	useState,
	type KeyboardEvent,
	type ReactNode,
	type RefObject
} from "react"
import { useTranslation } from "react-i18next"
import { useRouter } from "@tanstack/react-router"
import { defaultRangeExtractor, useVirtualizer, type Range, type Virtualizer } from "@tanstack/react-virtual"
import { SearchXIcon, SlidersHorizontalIcon, ClapperboardIcon } from "lucide-react"
import { useElementSize } from "@/lib/useElementSize"
import { observeElementOffsetFromAttach } from "@/lib/virtualScroll"
import { SearchInput } from "@/features/drive/components/searchInput"
import { columnsForWidth } from "@/features/drive/lib/gridLayout"
import { DEFAULT_DENSITY_INDEX, tileSizeForDensity } from "@/features/photos/lib/gridDensity"
import { useVideoHomeData } from "@/features/videos/hooks/useVideoHomeData"
import { useVideoCardActions } from "@/features/videos/hooks/useVideoCardActions"
import { useCardFocus } from "@/features/videos/hooks/useCardFocus"
import { useHistoryEntryKey } from "@/features/videos/hooks/useHistoryEntryKey"
import { recallScroll, rememberScroll } from "@/features/videos/lib/scrollMemory"
import { CardObserver } from "@/features/videos/lib/cardObserver"
import { useGridPrefsQuery, useProbeFactsQuery, useVideosGridDensityQuery, setGridPrefs } from "@/features/videos/queries/grid"
import { rowLiftPadding, titleCard, rowShape, type CardShape, type HomeData } from "@/features/videos/lib/homeRows.logic"
import {
	DEFAULT_GRID_PREFS,
	GRID_QUALITIES,
	GRID_SORTS,
	gridKeyAction,
	gridPrefsKey,
	withActiveRow,
	gridView,
	needsProbeFacts,
	sourceTitles,
	type GridPrefs,
	type GridQuality,
	type GridSort,
	type GridSource,
	type GridWatchFilter,
	type ProbeFacts
} from "@/features/videos/lib/videoGrid.logic"
import { filterTitlesByQuery, titleSearchTokens } from "@/features/videos/lib/search.logic"
import { VideoCard, type VideoCardAction } from "@/features/videos/components/videoCard"
import type { VideoCardModel } from "@/features/videos/lib/homeRows.logic"
import type { VideoTitleView } from "@/workers/videoIndex/view.logic"
import { Button } from "@/components/ui/button"
import { Select, SelectContent, SelectGroup, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select"
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover"
import { EmptyMessage } from "@/components/emptyMessage"

const GRID_GAP = 16
const GRID_OVERSCAN = 2
// Title and year under a card.
const CAPTION_HEIGHT = 46
const PHONE_WIDTH = 640
// A landscape card's target width against the density's poster width.
const LANDSCAPE_TILE_SCALE = 4 / 3

const SORT_KEYS = {
	title: "videosGridSortTitle",
	added: "videosGridSortAdded",
	year: "videosGridSortYear",
	runtime: "videosGridSortRuntime",
	played: "videosGridSortPlayed",
	random: "videosGridSortRandom"
} as const satisfies Record<GridSort, string>

const QUALITY_KEYS = {
	all: "videosGridQualityAll",
	"2160": "videosGridQuality2160",
	"1080": "videosGridQuality1080",
	"720": "videosGridQuality720",
	sd: "videosGridQualitySd"
} as const satisfies Record<GridQuality, string>

const WATCH_CHIPS: readonly {
	value: GridWatchFilter
	key: "videosGridFilterAll" | "videosGridFilterUnwatched" | "videosGridFilterInProgress" | "videosGridFilterWatched"
}[] = [
	{ value: "all", key: "videosGridFilterAll" },
	{ value: "unwatched", key: "videosGridFilterUnwatched" },
	{ value: "inProgress", key: "videosGridFilterInProgress" },
	{ value: "watched", key: "videosGridFilterWatched" }
]

const EMPTY_FACTS: ReadonlyMap<string, ProbeFacts> = new Map()

// One session seed for the Random sort, so it reads the same until the page reloads.
const RANDOM_SEED = Math.floor(Math.random() * 0x7fffffff)

function searchLabelKey(source: GridSource): "videosGridSearchMovies" | "videosGridSearchShows" | "videosGridSearchTitles" {
	return source.kind === "movies"
		? "videosGridSearchMovies"
		: source.kind === "shows"
			? "videosGridSearchShows"
			: "videosGridSearchTitles"
}

interface GridControlsProps {
	prefs: GridPrefs
	rootOptions: readonly { uuid: string; name: string }[]
	showQuality: boolean
	onChange: (prefs: GridPrefs) => void
}

// Sort, the watch-state chips and the resolution and directory selects: inline on wide panels, in a popover on phones.
function GridControls({ prefs, rootOptions, showQuality, onChange }: GridControlsProps) {
	const { t } = useTranslation("videos")
	const sortItems = GRID_SORTS.map(sort => ({ value: sort, label: t(SORT_KEYS[sort]) }))
	const qualityItems = GRID_QUALITIES.map(quality => ({ value: quality, label: t(QUALITY_KEYS[quality]) }))
	const rootItems = [{ value: "", label: t("videosGridRootAll") }, ...rootOptions.map(root => ({ value: root.uuid, label: root.name }))]

	return (
		<>
			<Select
				items={sortItems}
				value={prefs.sort}
				onValueChange={sort => {
					if (sort !== null) {
						onChange({ ...prefs, sort })
					}
				}}
			>
				<SelectTrigger
					size="sm"
					aria-label={t("videosGridSortLabel")}
					className="min-w-32"
				>
					<SelectValue />
				</SelectTrigger>
				<SelectContent>
					<SelectGroup>
						{sortItems.map(item => (
							<SelectItem
								key={item.value}
								value={item.value}
							>
								{item.label}
							</SelectItem>
						))}
					</SelectGroup>
				</SelectContent>
			</Select>
			<div
				role="group"
				aria-label={t("videosGridFilterLabel")}
				className="flex flex-wrap items-center gap-1"
			>
				{WATCH_CHIPS.map(chip => (
					<Button
						key={chip.value}
						variant={prefs.watch === chip.value ? "secondary" : "outline"}
						size="xs"
						className="rounded-full"
						aria-pressed={prefs.watch === chip.value}
						onClick={() => {
							onChange({ ...prefs, watch: chip.value })
						}}
					>
						{t(chip.key)}
					</Button>
				))}
			</div>
			{showQuality ? (
				<Select
					items={qualityItems}
					value={prefs.quality}
					onValueChange={quality => {
						if (quality !== null) {
							onChange({ ...prefs, quality })
						}
					}}
				>
					<SelectTrigger
						size="sm"
						aria-label={t("videosGridQualityLabel")}
					>
						<SelectValue />
					</SelectTrigger>
					<SelectContent>
						<SelectGroup>
							{qualityItems.map(item => (
								<SelectItem
									key={item.value}
									value={item.value}
								>
									{item.label}
								</SelectItem>
							))}
						</SelectGroup>
					</SelectContent>
				</Select>
			) : null}
			{rootOptions.length > 1 ? (
				<Select
					items={rootItems}
					value={prefs.root}
					onValueChange={root => {
						if (root !== null) {
							onChange({ ...prefs, root })
						}
					}}
				>
					<SelectTrigger
						size="sm"
						aria-label={t("videosGridRootLabel")}
						className="max-w-48"
					>
						<SelectValue />
					</SelectTrigger>
					<SelectContent>
						<SelectGroup>
							{rootItems.map(item => (
								<SelectItem
									key={item.value}
									value={item.value}
								>
									{item.label}
								</SelectItem>
							))}
						</SelectGroup>
					</SelectContent>
				</Select>
			) : null}
		</>
	)
}

// The grid's virtualizer as its rows component hands it over, for keyboard moves to scroll a row into range. A hook,
// so the grid never reads the ref while rendering.
function useRowScroller() {
	const ref = useRef<Virtualizer<HTMLDivElement, Element> | null>(null)

	return {
		ref,
		scrollToRow: (row: number): void => {
			ref.current?.scrollToIndex(row, { align: "auto" })
		}
	}
}

export interface VideoGridProps {
	source: GridSource
	// The view's name, for the list's accessible name.
	name: string
	// Library directories by uuid, for the directory select.
	rootNames: ReadonlyMap<string, string>
}

// The view's titles in order and what its cards are built from: a hook of its own, so the cursor and layout state of
// the grid never join their memoization.
function useGridModel({
	source,
	query,
	prefs,
	facts
}: {
	source: GridSource
	query: string
	prefs: GridPrefs
	facts: ReadonlyMap<string, ProbeFacts>
}) {
	const data = useVideoHomeData()
	const titles = sourceTitles(data.snapshot.titles, source)
	const tokens = titleSearchTokens(titles)
	const searched = filterTitlesByQuery({ titles, query, tokens })
	const view = gridView({ titles: searched, prefs, watch: data.watch, facts, seed: RANDOM_SEED })
	const shape: CardShape = source.kind === "other" ? "landscape" : source.kind === "root" ? rowShape(titles) : "poster"
	const homeData: Pick<HomeData, "index" | "entries" | "watch"> = { index: data.index, entries: data.entries, watch: data.watch }

	return { data, titles, view, shape, homeData }
}

// A library view: the Photos-style toolbar (search, sort, watch chips, resolution and directory, count, a coverage
// hint for sorts and filters only a probe can answer) over a virtualized grid of cards. Sort and filters are
// remembered per view.
export function VideoGrid({ source, name, rootNames }: VideoGridProps) {
	const { t } = useTranslation("videos")
	const prefsQuery = useGridPrefsQuery(source)
	const densityQuery = useVideosGridDensityQuery()
	const prefs = prefsQuery.data ?? DEFAULT_GRID_PREFS
	const factsQuery = useProbeFactsQuery(needsProbeFacts(prefs))
	const [query, setQuery] = useState("")
	const { data, titles, view, shape, homeData } = useGridModel({ source, query, prefs, facts: factsQuery.data ?? EMPTY_FACTS })
	const onAction = useVideoCardActions(data.items)
	const [scrollElement, setScrollElement] = useState<HTMLDivElement | null>(null)
	const { width } = useElementSize(scrollElement)
	const [activeId, setActiveId] = useState<string | null>(null)
	const rowScroller = useRowScroller()
	const cardFocus = useCardFocus<number>()
	const router = useRouter()
	const historyKey = useHistoryEntryKey()
	const scrollKey = `${historyKey}:grid:${gridPrefsKey(source)}`
	const [restoreOffset] = useState(() => recallScroll(scrollKey))

	// Cards within half a screen of the visible part may read and decode art; the overscan rows beyond wait.
	const nearObserver = useMemo(
		() => (scrollElement === null ? null : new CardObserver({ root: scrollElement, rootMargin: "50% 0px", threshold: 0 })),
		[scrollElement]
	)

	useEffect(() => {
		return () => {
			nearObserver?.disconnect()
		}
	}, [nearObserver])

	const contentWidth = width
	const phone = width > 0 && width < PHONE_WIDTH
	const tile = tileSizeForDensity(densityQuery.data ?? DEFAULT_DENSITY_INDEX) * (shape === "landscape" ? LANDSCAPE_TILE_SCALE : 1)
	const columns = phone ? (shape === "poster" ? 3 : 2) : columnsForWidth(contentWidth, tile, GRID_GAP)
	const cardWidth = columns > 0 ? (contentWidth - GRID_GAP * (columns - 1)) / columns : 0
	const artHeight = Math.round(cardWidth * (shape === "poster" ? 1.5 : 9 / 16))
	const rowHeight = artHeight + CAPTION_HEIGHT
	const rowCount = Math.ceil(view.titles.length / columns)
	const activeIndex = Math.max(0, activeId === null ? 0 : view.titles.findIndex(title => title.id === activeId))
	const rootOptions =
		source.kind === "root"
			? []
			: [...new Set(titles.flatMap(title => title.roots))].map(uuid => ({ uuid, name: rootNames.get(uuid) ?? "" }))
	const showQuality = source.kind !== "shows"
	const filtersActive = prefs.watch !== "all" || prefs.quality !== "all" || prefs.root !== "" || query.length > 0

	function updatePrefs(next: GridPrefs): void {
		setGridPrefs(source, next)
		setActiveId(null)
		scrollElement?.scrollTo({ top: 0 })
	}

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

		const action = gridKeyAction({
			key: event.key,
			index: activeIndex,
			count: view.titles.length,
			columns,
			pageRows: Math.max(1, Math.floor((scrollElement?.clientHeight ?? 0) / (rowHeight + GRID_GAP)))
		})

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
			const title = view.titles[activeIndex]

			if (title !== undefined) {
				onAction("play", titleCard({ data: homeData, title, shape }))
			}

			return
		}

		setActiveId(view.titles[action.index]?.id ?? null)
		rowScroller.scrollToRow(Math.floor(action.index / columns))
		cardFocus.focus(action.index, { scroll: false })
	}

	if (prefsQuery.status === "pending") {
		return null
	}

	const controls = (
		<GridControls
			prefs={prefs}
			rootOptions={rootOptions}
			showQuality={showQuality}
			onChange={updatePrefs}
		/>
	)

	let body: ReactNode

	if (titles.length === 0) {
		body = (
			<div className="flex flex-1 overflow-y-auto">
				<EmptyMessage
					icon={ClapperboardIcon}
					title={t("videosGridEmptyTitle")}
					description={t("videosGridEmptyBody")}
				/>
			</div>
		)
	} else if (view.titles.length === 0) {
		body = (
			<div className="flex flex-1 overflow-y-auto">
				<EmptyMessage
					icon={SearchXIcon}
					title={t("videosGridNoMatchesTitle")}
					description={t("videosGridNoMatchesBody")}
				>
					<Button
						variant="outline"
						onClick={() => {
							setQuery("")
							updatePrefs({ ...prefs, watch: "all", quality: "all", root: "" })
						}}
					>
						{t("videosGridClearFilters")}
					</Button>
				</EmptyMessage>
			</div>
		)
	} else {
		body = null
	}

	return (
		<div className="flex min-h-0 flex-1 flex-col overflow-hidden">
			<div className="flex shrink-0 flex-wrap items-center gap-x-3 gap-y-2 px-4 pb-3">
				<SearchInput
					action="videos.search"
					label={t(searchLabelKey(source))}
					value={query}
					onChange={next => {
						setQuery(next)
						setActiveId(null)
					}}
					onClear={() => {
						setQuery("")
					}}
				/>
				<div className="flex flex-wrap items-center gap-x-3 gap-y-2 max-sm:hidden">{controls}</div>
				<Popover>
					<PopoverTrigger
						render={
							<Button
								variant={filtersActive ? "secondary" : "outline"}
								size="sm"
								className="sm:hidden"
							>
								<SlidersHorizontalIcon data-icon="inline-start" />
								{t("videosGridFilters")}
							</Button>
						}
					/>
					<PopoverContent
						align="end"
						className="flex w-72 flex-col items-start gap-3"
					>
						{controls}
					</PopoverContent>
				</Popover>
				<span
					role="status"
					className="ml-auto text-xs text-muted-foreground tabular-nums"
				>
					{t("videosGridCount", { count: view.titles.length })}
				</span>
			</div>
			{view.coverage !== null && view.coverage.known < view.coverage.total ? (
				<p className="shrink-0 px-4 pb-2 text-xs text-muted-foreground tabular-nums">
					{view.coverage.kind === "runtime"
						? t("videosGridCoverageRuntime", { known: view.coverage.known, total: view.coverage.total })
						: t("videosGridCoverageQuality", { known: view.coverage.known, total: view.coverage.total })}
				</p>
			) : null}
			{body ?? (
				<div
					ref={setScrollElement}
					role="list"
					aria-label={t("videosGridLabel", { name })}
					className="min-h-0 flex-1 overflow-y-auto px-4 pb-6"
					style={{ paddingTop: rowLiftPadding(artHeight) }}
					onKeyDown={handleKeyDown}
					onScroll={event => {
						rememberScroll(scrollKey, event.currentTarget.scrollTop)
					}}
				>
					{width > 0 ? (
						<VideoGridRows
							scrollElement={scrollElement}
							titles={view.titles}
							rowCount={rowCount}
							columns={columns}
							rowHeight={rowHeight}
							cardWidth={cardWidth}
							shape={shape}
							homeData={homeData}
							activeIndex={activeIndex}
							nearObserver={nearObserver}
							restoreOffset={restoreOffset}
							registerCard={cardFocus.register}
							virtualizerRef={rowScroller.ref}
							onFocusCard={index => {
								setActiveId(view.titles[index]?.id ?? null)
							}}
							onAction={onAction}
						/>
					) : null}
				</div>
			)}
		</div>
	)
}

interface GridCardProps {
	title: VideoTitleView
	shape: CardShape
	homeData: Pick<HomeData, "index" | "entries" | "watch">
	index: number
	setSize: number
	active: boolean
	nearObserver: CardObserver | null
	registerCard: (index: number, element: HTMLAnchorElement | null) => void
	onFocusCard: (index: number) => void
	onAction: (action: VideoCardAction, card: VideoCardModel) => void
}

// Compiled, unlike the virtualizer host around it: builds the card's model and callbacks from what it is given, so a
// scroll step or cursor move leaves every card whose props didn't change as it was.
function GridCard({ title, shape, homeData, index, setSize, active, nearObserver, registerCard, onFocusCard, onAction }: GridCardProps) {
	return (
		<VideoCard
			card={titleCard({ data: homeData, title, shape })}
			nearObserver={nearObserver}
			active={active}
			posInSet={index + 1}
			setSize={setSize}
			registerRef={element => {
				registerCard(index, element)
			}}
			onFocusCard={() => {
				onFocusCard(index)
			}}
			onAction={onAction}
		/>
	)
}

interface VideoGridRowsProps {
	scrollElement: HTMLDivElement | null
	titles: readonly VideoTitleView[]
	rowCount: number
	columns: number
	rowHeight: number
	cardWidth: number
	shape: CardShape
	homeData: Pick<HomeData, "index" | "entries" | "watch">
	activeIndex: number
	nearObserver: CardObserver | null
	// Where Back left the list, applied once the rows are laid out.
	restoreOffset: number | undefined
	registerCard: (index: number, element: HTMLAnchorElement | null) => void
	virtualizerRef: RefObject<Virtualizer<HTMLDivElement, Element> | null>
	onFocusCard: (index: number) => void
	onAction: (action: VideoCardAction, card: VideoCardModel) => void
}

// Owns the virtualizer, which opts its host out of the React Compiler and re-renders it on every range change; split
// out so VideoGrid compiles. The row holding the cursor stays mounted when scrolled away, so focus never drops.
function VideoGridRows({
	scrollElement,
	titles,
	rowCount,
	columns,
	rowHeight,
	cardWidth,
	shape,
	homeData,
	activeIndex,
	nearObserver,
	restoreOffset,
	registerCard,
	virtualizerRef,
	onFocusCard,
	onAction
}: VideoGridRowsProps) {
	// Changing with the layout is what makes the virtualizer re-read row sizes; stable while scrolling.
	const getRowKey = useCallback((index: number) => `${String(columns)}:${String(rowHeight)}:${String(index)}`, [columns, rowHeight])
	const activeRow = Math.floor(activeIndex / columns)
	const rangeExtractor = useCallback((range: Range) => withActiveRow(defaultRangeExtractor(range), activeRow, range.count), [activeRow])
	const virtualizer = useVirtualizer({
		count: rowCount,
		getScrollElement: () => scrollElement,
		observeElementOffset: observeElementOffsetFromAttach,
		estimateSize: () => rowHeight,
		gap: GRID_GAP,
		overscan: GRID_OVERSCAN,
		getItemKey: getRowKey,
		rangeExtractor
	})
	const restored = useRef(false)

	useLayoutEffect(() => {
		virtualizerRef.current = virtualizer

		return () => {
			virtualizerRef.current = null
		}
	}, [virtualizer, virtualizerRef])

	const totalSize = virtualizer.getTotalSize()

	useLayoutEffect(() => {
		if (restored.current || restoreOffset === undefined || scrollElement === null || totalSize === 0) {
			return
		}

		restored.current = true
		scrollElement.scrollTo({ top: restoreOffset })
	}, [restoreOffset, scrollElement, totalSize])

	return (
		<div style={{ position: "relative", width: "100%", height: totalSize }}>
			{virtualizer.getVirtualItems().map(virtualRow => {
				const start = virtualRow.index * columns

				return (
					<div
						key={virtualRow.key}
						style={{
							position: "absolute",
							top: 0,
							left: 0,
							width: "100%",
							height: virtualRow.size,
							transform: `translateY(${String(virtualRow.start)}px)`,
							display: "grid",
							gridTemplateColumns: `repeat(${String(columns)}, ${String(cardWidth)}px)`,
							columnGap: GRID_GAP
						}}
					>
						{titles.slice(start, start + columns).map((title, column) => {
							const index = start + column

							return (
								<GridCard
									key={title.id}
									title={title}
									shape={shape}
									homeData={homeData}
									index={index}
									setSize={titles.length}
									active={index === activeIndex}
									nearObserver={nearObserver}
									registerCard={registerCard}
									onFocusCard={onFocusCard}
									onAction={onAction}
								/>
							)
						})}
					</div>
				)
			})}
		</div>
	)
}
