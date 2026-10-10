import { useState } from "react"
import { useTranslation } from "react-i18next"
import { useNavigate, useRouter } from "@tanstack/react-router"
import { ClapperboardIcon } from "lucide-react"
import { sdkApi } from "@/lib/sdk/client"
import { videoLibrary } from "@/features/videos/lib/index"
import { useVideoLibrary, useVideoTitleDetailQuery } from "@/features/videos/queries/library"
import { useCatalogSeriesQuery, useCatalogTitleQuery } from "@/features/videos/queries/catalog"
import { useHistoryEntryKey } from "@/features/videos/hooks/useHistoryEntryKey"
import { recallScroll, rememberScroll } from "@/features/videos/lib/scrollMemory"
import { titleSubject, type CorrectionSubject } from "@/features/videos/lib/corrections.logic"
import { detailHeaderMode } from "@/features/videos/lib/titleDetail.logic"
import { probeLandscapeCover } from "@/features/videos/lib/artwork.logic"
import { useStoredProbesQuery } from "@/features/videos/queries/detail"
import { MovieDetail } from "@/features/videos/components/detail/movieDetail"
import { SeriesDetail } from "@/features/videos/components/detail/seriesDetail"
import { CorrectionsDialog } from "@/features/videos/components/detail/correctionsDialog"
import { LoadingState } from "@/components/loadingState"
import { Empty, EmptyContent, EmptyDescription, EmptyHeader, EmptyMedia, EmptyTitle } from "@/components/ui/empty"
import { Button } from "@/components/ui/button"

// After a correction the item may be shown as another title; the page follows once the index has it (a
// classification is local and quick), and stays where it is if it never turns up.
const FOLLOW_TIMEOUT_MS = 5_000

// Asks the library's worker again each time its catalog changes, until the title is there or the wait is over.
function whenTitleKnown(id: string, then: () => void): void {
	let done = false
	let version = -1
	let unsubscribe: () => void = () => undefined
	const finish = (): void => {
		done = true
		clearTimeout(timer)
		unsubscribe()
	}
	const timer = setTimeout(finish, FOLLOW_TIMEOUT_MS)
	const check = (): void => {
		const current = videoLibrary.getSnapshot().version

		if (done || current === version) {
			return
		}

		version = current
		void sdkApi.videoCatalogTitle(id).then(title => {
			if (title !== null && !done) {
				finish()
				then()
			}
		})
	}

	unsubscribe = videoLibrary.subscribe(check)
	check()
}

// /videos/title/$uuid: a movie's, another video's or a series' page, from the local index (no listing).
export function VideoTitleScreen({ titleId }: { titleId: string }) {
	const { t } = useTranslation("videos")
	const navigate = useNavigate()
	const router = useRouter()
	const library = useVideoLibrary()
	const titleQuery = useCatalogTitleQuery(titleId)
	const title = titleQuery.data ?? null
	const scanning = library.roots.some(root => root.walking || root.state === "loading")
	const detailQuery = useVideoTitleDetailQuery(titleId)
	const detail = detailQuery.data ?? null
	const [editing, setEditing] = useState<CorrectionSubject | null>(null)
	const seriesQuery = useCatalogSeriesQuery(editing !== null)
	// A Matroska movie without a backdrop image may carry a landscape cover its stored probe located (a local read,
	// waited for so the header never switches form after it shows).
	const primary = title?.primary ?? null
	const coverCheck = title !== null && title.backdrop === null && title.kind === "movie" && title.container === "mkv" ? primary : null
	const storedCover = useStoredProbesQuery(coverCheck === null ? [] : [coverCheck.uuid])
	const headerMode =
		title === null
			? "band"
			: detailHeaderMode(title, coverCheck !== null && probeLandscapeCover(storedCover.data?.get(coverCheck.uuid)) !== null)
	const historyKey = useHistoryEntryKey()
	const scrollKey = `${historyKey}:title:${titleId}`
	const [restoreOffset] = useState(() => recallScroll(scrollKey))

	function back(): void {
		if (router.history.canGoBack()) {
			router.history.back()

			return
		}

		void navigate({ to: title?.kind === "series" ? "/videos/shows" : title?.kind === "other" ? "/videos/other" : "/videos/movies" })
	}

	if (
		!library.loaded ||
		titleQuery.isPending ||
		(title === null && scanning) ||
		(title !== null && detailQuery.isPending) ||
		(coverCheck !== null && storedCover.isPending)
	) {
		return <LoadingState size="lg" />
	}

	if (title === null || detail === null) {
		return (
			<div className="flex min-h-0 flex-1 items-center justify-center p-6">
				<Empty className="max-w-md border border-dashed">
					<EmptyHeader>
						<EmptyMedia>
							<ClapperboardIcon />
						</EmptyMedia>
						<EmptyTitle>{t("detailNotFound")}</EmptyTitle>
						<EmptyDescription>{t("detailNotFoundDescription")}</EmptyDescription>
					</EmptyHeader>
					<EmptyContent>
						<Button
							variant="outline"
							onClick={back}
						>
							{t("detailBack")}
						</Button>
					</EmptyContent>
				</Empty>
			</div>
		)
	}

	const series = seriesQuery.data ?? []

	return (
		<div
			ref={element => {
				if (element !== null && restoreOffset !== undefined && element.scrollTop === 0) {
					element.scrollTop = restoreOffset
				}
			}}
			className="min-h-0 flex-1 overflow-y-auto"
			onScroll={event => {
				rememberScroll(scrollKey, event.currentTarget.scrollTop)
			}}
		>
			{detail.kind === "series" ? (
				<SeriesDetail
					key={detail.title.id}
					detail={detail}
					headerMode={headerMode}
					onBack={back}
					onEdit={setEditing}
				/>
			) : (
				<MovieDetail
					key={detail.title.id}
					detail={detail}
					headerMode={headerMode}
					onBack={back}
					onEdit={() => {
						setEditing(titleSubject(detail))
					}}
				/>
			)}
			{editing === null ? null : (
				<CorrectionsDialog
					key={editing.key}
					subject={editing}
					series={series}
					onClose={() => {
						setEditing(null)
					}}
					onMoved={id => {
						if (id === titleId) {
							return
						}

						whenTitleKnown(id, () => {
							void navigate({ to: "/videos/title/$uuid", params: { uuid: id }, replace: true })
						})
					}}
				/>
			)}
		</div>
	)
}
