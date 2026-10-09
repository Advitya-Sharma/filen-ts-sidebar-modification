import { useEffect, useMemo, useState } from "react"
import { useTranslation } from "react-i18next"
import { Link } from "@tanstack/react-router"
import type { ExtraType } from "@filen/shared/videos"
import { cn } from "@filen/shared"
import { useVideoArt } from "@/features/videos/hooks/useVideoArt"
import { CardObserver } from "@/features/videos/lib/cardObserver"
import { extrasShelf } from "@/features/videos/lib/titleDetail.logic"
import { ArtImage } from "@/features/videos/components/videoCard"
import type { VideoExtraView } from "@/workers/videoIndex/view.logic"
import type { VideosKey } from "@/lib/i18n"

// The detail pages' extras: trailers first, then featurettes and the rest, as 16:9 stills that load only once the
// shelf comes near the viewport. Each opens the player on that extra.

const EXTRA_TYPE_KEY: Readonly<Record<ExtraType, VideosKey>> = {
	trailer: "detailExtraTrailer",
	featurette: "detailExtraFeaturette",
	behindTheScenes: "detailExtraBehindTheScenes",
	deletedScene: "detailExtraDeletedScene",
	interview: "detailExtraInterview",
	scene: "detailExtraScene",
	short: "detailExtraShort",
	clip: "detailExtraClip",
	unknown: "detailExtraUnknown"
}

function ExtraCard({ titleId, extra, observer }: { titleId: string; extra: VideoExtraView; observer: CardObserver | null }) {
	const { t } = useTranslation("videos")
	const [element, setElement] = useState<HTMLAnchorElement | null>(null)
	const [near, setNear] = useState(false)
	const stableUuid = extra.part.file.stableUuid

	useEffect(() => {
		if (element === null || observer === null) {
			return
		}

		return observer.observe(element, inView => {
			if (inView) {
				setNear(true)
			}
		})
	}, [element, observer])

	const art = useVideoArt({ titleId, stableUuid, slot: "s", title: extra.title, near, element })
	const type = t(EXTRA_TYPE_KEY[extra.type])

	return (
		<li className="w-56 shrink-0 snap-start">
			<Link
				ref={setElement}
				to="/videos/watch/$uuid"
				params={{ uuid: titleId }}
				search={{ file: stableUuid }}
				aria-label={t("detailExtraLabel", { type, title: extra.title })}
				className="group/extra block rounded-xl outline-none"
			>
				<div
					className={cn(
						"relative aspect-video w-full overflow-hidden rounded-xl bg-muted/40",
						"group-focus-visible/extra:ring-2 group-focus-visible/extra:ring-ring",
						extra.part.unsupported && "opacity-60"
					)}
				>
					{art.kind === "image" ? (
						<ArtImage
							key={art.url}
							url={art.url}
							fit={art.fit}
						/>
					) : null}
				</div>
				<p
					aria-hidden="true"
					className="mt-2 truncate text-sm font-medium"
				>
					{extra.title}
				</p>
				<p
					aria-hidden="true"
					className="truncate text-xs text-muted-foreground"
				>
					{type}
				</p>
			</Link>
		</li>
	)
}

export function ExtrasShelf({ titleId, extras, headingId }: { titleId: string; extras: readonly VideoExtraView[]; headingId: string }) {
	const { t } = useTranslation("videos")
	const [scroller, setScroller] = useState<HTMLUListElement | null>(null)
	// Stills within a screen of the visible part, either way.
	const observer = useMemo(
		() => (scroller === null ? null : new CardObserver({ root: null, rootMargin: "50% 50%", threshold: 0 })),
		[scroller]
	)

	useEffect(() => {
		return () => {
			observer?.disconnect()
		}
	}, [observer])

	const shelf = extrasShelf(extras)

	if (shelf.length === 0) {
		return null
	}

	return (
		<section
			aria-labelledby={headingId}
			className="flex flex-col gap-3"
		>
			<h2
				id={headingId}
				className="text-sm font-medium"
			>
				{t("detailExtras")}
			</h2>
			<ul
				ref={setScroller}
				className="-mx-4 flex snap-x scroll-px-4 gap-3 overflow-x-auto px-4 pb-2 md:-mx-6 md:scroll-px-6 md:px-6"
			>
				{shelf.map(extra => (
					<ExtraCard
						key={extra.part.file.stableUuid}
						titleId={titleId}
						extra={extra}
						observer={observer}
					/>
				))}
			</ul>
		</section>
	)
}
