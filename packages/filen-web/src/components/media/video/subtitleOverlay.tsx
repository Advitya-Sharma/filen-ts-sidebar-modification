import { useEffect, useRef, useState, type CSSProperties } from "react"
import type { CueStore } from "@filen/shared/videos"
import { cn } from "@filen/shared"
import { SubtitleTrackPool } from "@/lib/media/subtitles/subtitleTracks"
import {
	cueSegments,
	subtitleCss,
	subtitleRegionsAt,
	subtitleRegionsKey,
	type SubtitleRegion
} from "@/lib/media/subtitles/subtitleStyle.logic"
import { useSubtitleStyleStore, warmSubtitleStyle } from "@/lib/media/subtitles/subtitleStyle"
import { isVideoNativeFullscreen } from "@/lib/media/fullscreen"

// Subtitles over the picture. The cues live in one of the element's own text tracks (hidden); its
// cuechange drives this overlay, which re-renders only when what shows changes. A change of cues, delay or
// style while paused redraws at once (no cuechange fires then). In picture-in-picture and iPhone's native
// fullscreen the browser draws the track itself instead.

// One pool per element, for the element's life (its tracks can't be removed anyway).
const pools = new WeakMap<HTMLVideoElement, SubtitleTrackPool>()

function trackPool(video: HTMLVideoElement): SubtitleTrackPool {
	let pool = pools.get(video)

	if (pool === undefined) {
		pool = new SubtitleTrackPool(video)
		pools.set(video, pool)
	}

	return pool
}

// The track's cues are rebuilt this long after the last delay change (the overlay follows at once).
const CUE_REBUILD_DELAY_MS = 300

const REGION_CLASS: Readonly<Record<SubtitleRegion["key"], string>> = {
	bottom: "bottom-(--subtitle-bottom)",
	middle: "top-1/2 -translate-y-1/2",
	top: "top-[6%]"
}

const ALIGN_CLASS: Readonly<Record<SubtitleRegion["align"], string>> = {
	left: "items-start text-left",
	center: "items-center text-center",
	right: "items-end text-right"
}

export function SubtitleOverlay({
	video,
	cues,
	offset,
	sample,
	raised
}: {
	video: HTMLVideoElement | null
	cues: CueStore | null
	// Seconds, positive = later.
	offset: number
	// Shown instead of the cues (the style panel's live preview).
	sample?: string | null | undefined
	// The controls are up: bottom lines move above the bar (top and middle ones stay put).
	raised?: boolean | undefined
}) {
	const style = useSubtitleStyleStore()
	const css = subtitleCss(style)
	const [regions, setRegions] = useState<SubtitleRegion[]>([])
	const [native, setNative] = useState(false)
	const pool = video === null ? null : trackPool(video)
	const shownKey = useRef("")
	const scratch = useRef<number[]>([])

	useEffect(() => {
		void warmSubtitleStyle()
	}, [])

	useEffect(
		() => () => {
			pool?.dispose()
		},
		[pool]
	)

	// The track: rebuilt when the cues change, and shortly after the delay settles.
	useEffect(() => {
		if (pool === null) {
			return
		}

		const timer = setTimeout(
			() => {
				pool.show(cues, offset)
			},
			cues === null ? 0 : CUE_REBUILD_DELAY_MS
		)

		return () => {
			clearTimeout(timer)
		}
	}, [pool, cues, offset])

	// The overlay: drawn from the cue store at the element's time, on every cuechange and seek.
	useEffect(() => {
		if (video === null || pool === null) {
			return
		}

		const element = video

		function draw(): void {
			const next = cues === null ? [] : subtitleRegionsAt(cues, element.currentTime, offset, scratch.current)
			const key = subtitleRegionsKey(next)

			if (key !== shownKey.current) {
				shownKey.current = key
				setRegions(next)
			}
		}

		draw()
		pool.onCueChange(draw)
		element.addEventListener("seeked", draw)
		element.addEventListener("emptied", draw)

		return () => {
			pool.onCueChange(null)
			element.removeEventListener("seeked", draw)
			element.removeEventListener("emptied", draw)
		}
	}, [video, pool, cues, offset])

	// The browser draws the track where the overlay can't be seen.
	useEffect(() => {
		if (video === null || pool === null) {
			return
		}

		const element = video

		function update(): void {
			const next = document.pictureInPictureElement === element || isVideoNativeFullscreen(element)

			setNative(next)
			pool?.setNative(next)
		}

		update()

		const events = ["enterpictureinpicture", "leavepictureinpicture", "webkitbeginfullscreen", "webkitendfullscreen"]

		for (const event of events) {
			element.addEventListener(event, update)
		}

		return () => {
			for (const event of events) {
				element.removeEventListener(event, update)
			}
		}
	}, [video, pool])

	const shown: SubtitleRegion[] =
		sample !== null && sample !== undefined && sample.length > 0
			? [{ key: "bottom", align: "center", cues: [sample] }]
			: native
				? []
				: regions

	if (shown.length === 0) {
		return null
	}

	return (
		<div
			aria-hidden="true"
			className="[container-type:size] pointer-events-none absolute inset-0"
			style={{ "--subtitle-bottom": `${String(css.bottom)}%` } as CSSProperties}
		>
			{shown.map(region => (
				<div
					key={`${region.key}:${region.align}`}
					className={cn(
						"absolute inset-x-[5%] flex flex-col gap-[0.15em]",
						REGION_CLASS[region.key],
						ALIGN_CLASS[region.align],
						region.key === "bottom" && "transition-transform duration-200 motion-reduce:transition-none",
						region.key === "bottom" && raised === true && "-translate-y-16"
					)}
					style={{
						fontFamily: css.fontFamily,
						fontSize: `max(14px, calc(${String(css.scale)} * 3.6cqh))`,
						lineHeight: 1.25
					}}
				>
					{region.cues.map((text, index) => (
						<span
							key={index}
							className="max-w-full rounded-[0.15em] px-[0.25em] font-medium whitespace-pre-line"
							style={{ color: css.color, background: css.background, textShadow: css.textShadow }}
						>
							{cueSegments(text).map((segment, at) => (
								<span
									key={at}
									className={cn(
										segment.italic && "italic",
										segment.bold && "font-bold",
										segment.underline && "underline"
									)}
								>
									{segment.text}
								</span>
							))}
						</span>
					))}
				</div>
			))}
		</div>
	)
}
