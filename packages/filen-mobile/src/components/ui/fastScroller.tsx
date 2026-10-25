import { useRef, useState, useEffect } from "react"
import { StyleSheet, View as RNView, type LayoutChangeEvent, type NativeScrollEvent, type NativeSyntheticEvent } from "react-native"
import { GestureDetector, usePanGesture } from "react-native-gesture-handler"
import Animated, {
	useSharedValue,
	useAnimatedStyle,
	useAnimatedReaction,
	withTiming,
	withDelay,
	interpolate,
	Extrapolation,
	type SharedValue
} from "react-native-reanimated"
import { scheduleOnRN } from "react-native-worklets"
import Ionicons from "@expo/vector-icons/Ionicons"
import Text from "@/components/ui/text"
import haptics from "@/lib/haptics"
import { useResolveClassNames } from "@/hooks/useResolveClassNames"

/**
 * Fast scroller ("scrubber") for long lists.
 *
 * - A draggable handle on the right edge: drag it to jump anywhere in the list instantly, no flinging.
 * - A bubble next to the handle names where you are ("Mar 2021" in Photos, "K" in a name-sorted folder).
 * - Optional markers (years, letters) are drawn along the track while dragging, so you can aim for one.
 * - The handle fades in while the list scrolls, fades out when idle, and is never shown for short lists.
 *   While hidden it is not touchable, so it never steals taps from the items underneath it.
 */

const THUMB_HEIGHT = 48
const THUMB_WIDTH = 28
const BUBBLE_HEIGHT = 36
const TRACK_PADDING_TOP = 8
const TRACK_PADDING_BOTTOM = 120
const IDLE_HIDE_DELAY_MS = 1500
const MIN_SCREENS_TO_SHOW = 3
const MIN_MARKER_GAP_PX = 26

export type FastScrollMarker = {
	label: string
	/** Index into the list's data where this marker starts. */
	index: number
}

export type FastScrollOptions = {
	/** Bubble text for the item at a data index. Return null to hide the bubble. */
	getLabel?: (index: number) => string | null
	markers?: FastScrollMarker[]
}

export type FastScrollController = {
	scrollY: SharedValue<number>
	contentHeight: SharedValue<number>
	viewportHeight: SharedValue<number>
	opacity: SharedValue<number>
	dragging: SharedValue<number>
	contentHeightJS: number
	viewportHeightJS: number
	onScroll: (e: NativeSyntheticEvent<NativeScrollEvent>) => void
	onContentSizeChange: (w: number, h: number) => void
	onLayout: (e: LayoutChangeEvent) => void
}

export function useFastScrollController(): FastScrollController {
	const scrollY = useSharedValue<number>(0)
	const contentHeight = useSharedValue<number>(0)
	const viewportHeight = useSharedValue<number>(0)
	const opacity = useSharedValue<number>(0)
	const dragging = useSharedValue<number>(0)
	const [contentHeightJS, setContentHeightJS] = useState<number>(0)
	const [viewportHeightJS, setViewportHeightJS] = useState<number>(0)

	return {
		scrollY,
		contentHeight,
		viewportHeight,
		opacity,
		dragging,
		contentHeightJS,
		viewportHeightJS,
		onScroll: e => {
			const { contentOffset, contentSize, layoutMeasurement } = e.nativeEvent

			scrollY.set(contentOffset.y)
			contentHeight.set(contentSize.height)
			viewportHeight.set(layoutMeasurement.height)

			if (dragging.get() === 0) {
				// Show while scrolling; fade away after a short idle period.
				opacity.set(
					withTiming(1, { duration: 120 }, finished => {
						"worklet"

						if (finished && dragging.value === 0) {
							opacity.value = withDelay(IDLE_HIDE_DELAY_MS, withTiming(0, { duration: 300 }))
						}
					})
				)
			}
		},
		onContentSizeChange: (_w, h) => {
			contentHeight.set(h)
			setContentHeightJS(prev => (Math.abs(prev - h) < 1 ? prev : h))
		},
		onLayout: e => {
			const h = e.nativeEvent.layout.height

			viewportHeight.set(h)
			setViewportHeightJS(prev => (Math.abs(prev - h) < 1 ? prev : h))
		}
	}
}

/** Pure: which markers fit on a track of `trackHeight` px without overlapping. Exported for tests. */
export function layoutFastScrollMarkers({
	markers,
	itemCount,
	trackHeight
}: {
	markers: FastScrollMarker[] | undefined
	itemCount: number
	trackHeight: number
}): { label: string; top: number }[] {
	if (!markers || markers.length === 0 || itemCount <= 1 || trackHeight <= 0) {
		return []
	}

	const result: { label: string; top: number }[] = []
	let lastTop = -Infinity

	for (const marker of markers) {
		const top = TRACK_PADDING_TOP + (marker.index / (itemCount - 1)) * trackHeight + THUMB_HEIGHT / 2 - 9

		if (top - lastTop >= MIN_MARKER_GAP_PX) {
			result.push({ label: marker.label, top })

			lastTop = top
		}
	}

	return result
}

type Scrollable = {
	scrollToOffset: (params: { offset: number; animated?: boolean }) => void
}

export const FastScroller = ({
	controller,
	getList,
	itemCount,
	options
}: {
	controller: FastScrollController
	getList: () => Scrollable | null
	itemCount: number
	options: FastScrollOptions
}) => {
	const { scrollY, contentHeight, viewportHeight, opacity, dragging, contentHeightJS, viewportHeightJS } = controller
	const primary = useResolveClassNames("bg-primary").backgroundColor as string
	const card = useResolveClassNames("bg-card").backgroundColor as string
	const foreground = useResolveClassNames("text-foreground").color as string
	const trackHeight = useSharedValue<number>(0)
	const dragStartTop = useSharedValue<number>(0)
	const dragProgress = useSharedValue<number>(0)
	const [trackHeightJS, setTrackHeightJS] = useState<number>(0)
	const [label, setLabel] = useState<string | null>(null)
	const [interactive, setInteractive] = useState<boolean>(false)
	const pendingProgress = useRef<number | null>(null)
	const rafId = useRef<number | null>(null)
	const lastLabel = useRef<string | null>(null)
	const lastMarker = useRef<number>(-1)

	const enabled = itemCount > 0 && viewportHeightJS > 0 && contentHeightJS > viewportHeightJS * MIN_SCREENS_TO_SHOW

	useAnimatedReaction(
		() => opacity.value > 0.05,
		(visible, previous) => {
			if (visible !== previous) {
				scheduleOnRN(setInteractive, visible)
			}
		}
	)

	useEffect(() => {
		return () => {
			if (rafId.current !== null) {
				cancelAnimationFrame(rafId.current)
			}
		}
	}, [])

	const indexForProgress = (progress: number) => Math.min(itemCount - 1, Math.max(0, Math.round(progress * (itemCount - 1))))

	const updateLabel = (progress: number) => {
		if (itemCount === 0) {
			return
		}

		const index = indexForProgress(progress)

		if (options.getLabel) {
			const next = options.getLabel(index)

			if (next !== lastLabel.current) {
				lastLabel.current = next

				setLabel(next)
			}
		}

		// A light tick each time the handle crosses a marker (a new year / letter).
		if (options.markers && options.markers.length > 0) {
			let markerIndex = -1

			for (let i = 0; i < options.markers.length; i++) {
				const marker = options.markers[i]

				if (marker && marker.index <= index) {
					markerIndex = i
				} else {
					break
				}
			}

			if (markerIndex !== lastMarker.current) {
				if (lastMarker.current !== -1) {
					haptics.selection()
				}

				lastMarker.current = markerIndex
			}
		}
	}

	// Coalesce drag updates into one scroll per frame.
	const flush = () => {
		rafId.current = null

		const progress = pendingProgress.current

		if (progress === null) {
			return
		}

		pendingProgress.current = null

		const maxOffset = Math.max(0, contentHeight.get() - viewportHeight.get())

		getList()?.scrollToOffset({
			offset: progress * maxOffset,
			animated: false
		})

		updateLabel(progress)
	}

	const scrollToProgress = (progress: number) => {
		pendingProgress.current = progress

		if (rafId.current === null) {
			rafId.current = requestAnimationFrame(flush)
		}
	}

	const onDragStart = (progress: number) => {
		lastLabel.current = null
		lastMarker.current = -1

		updateLabel(progress)
		haptics.selection()
	}

	const gesture = usePanGesture({
		minDistance: 0,
		hitSlop: { left: 24, right: 8, top: 12, bottom: 12 },
		onActivate: () => {
			"worklet"

			const maxOffset = Math.max(1, contentHeight.value - viewportHeight.value)
			const progress = Math.min(1, Math.max(0, scrollY.value / maxOffset))

			dragging.set(1)
			opacity.set(withTiming(1, { duration: 80 }))
			dragStartTop.value = progress * trackHeight.value
			dragProgress.value = progress

			scheduleOnRN(onDragStart, progress)
		},
		onUpdate: e => {
			"worklet"

			if (trackHeight.value <= 0) {
				return
			}

			const top = Math.min(trackHeight.value, Math.max(0, dragStartTop.value + e.translationY))
			const progress = top / trackHeight.value

			dragProgress.value = progress

			scheduleOnRN(scrollToProgress, progress)
		},
		onFinalize: () => {
			"worklet"

			dragging.set(0)
			opacity.set(withDelay(IDLE_HIDE_DELAY_MS, withTiming(0, { duration: 300 })))
		}
	})

	const thumbStyle = useAnimatedStyle(() => {
		const maxOffset = Math.max(1, contentHeight.value - viewportHeight.value)
		const progress = dragging.value === 1 ? dragProgress.value : Math.min(1, Math.max(0, scrollY.value / maxOffset))

		return {
			opacity: opacity.value,
			transform: [{ translateY: TRACK_PADDING_TOP + progress * trackHeight.value }]
		}
	})

	const bubbleStyle = useAnimatedStyle(() => {
		return {
			opacity: withTiming(dragging.value, { duration: 120 }),
			transform: [
				{ translateY: TRACK_PADDING_TOP + dragProgress.value * trackHeight.value + (THUMB_HEIGHT - BUBBLE_HEIGHT) / 2 },
				{ scale: interpolate(dragging.value, [0, 1], [0.9, 1], Extrapolation.CLAMP) }
			]
		}
	})

	const markersStyle = useAnimatedStyle(() => {
		return {
			opacity: withTiming(dragging.value, { duration: 150 })
		}
	})

	if (!enabled) {
		return null
	}

	const visibleMarkers = layoutFastScrollMarkers({
		markers: options.markers,
		itemCount,
		trackHeight: trackHeightJS
	})

	return (
		<RNView
			style={styles.container}
			pointerEvents="box-none"
			onLayout={e => {
				const h = Math.max(0, e.nativeEvent.layout.height - TRACK_PADDING_TOP - TRACK_PADDING_BOTTOM - THUMB_HEIGHT)

				trackHeight.set(h)
				setTrackHeightJS(h)
			}}
		>
			{visibleMarkers.length > 0 && (
				<Animated.View
					style={[StyleSheet.absoluteFill, markersStyle]}
					pointerEvents="none"
				>
					{visibleMarkers.map(marker => (
						<RNView
							key={`${marker.label}-${marker.top}`}
							style={[styles.marker, { top: marker.top, backgroundColor: card }]}
						>
							<Text
								className="text-xs font-semibold"
								style={{ color: foreground }}
							>
								{marker.label}
							</Text>
						</RNView>
					))}
				</Animated.View>
			)}
			{label !== null && (
				<Animated.View
					style={[styles.bubble, bubbleStyle, { backgroundColor: primary }]}
					pointerEvents="none"
				>
					<Text
						className="text-base font-bold"
						style={{ color: "#ffffff" }}
						numberOfLines={1}
					>
						{label}
					</Text>
				</Animated.View>
			)}
			<GestureDetector gesture={gesture}>
				<Animated.View
					pointerEvents={interactive ? "auto" : "none"}
					style={[styles.thumb, thumbStyle, { backgroundColor: card }]}
				>
					<Ionicons
						name="code-outline"
						size={18}
						color={foreground}
						style={styles.thumbIcon}
					/>
				</Animated.View>
			</GestureDetector>
		</RNView>
	)
}

const styles = StyleSheet.create({
	container: {
		position: "absolute",
		top: 0,
		bottom: 0,
		right: 0,
		width: 220,
		zIndex: 1000
	},
	thumb: {
		position: "absolute",
		right: 4,
		top: 0,
		width: THUMB_WIDTH,
		height: THUMB_HEIGHT,
		borderRadius: THUMB_WIDTH / 2,
		alignItems: "center",
		justifyContent: "center",
		elevation: 4,
		shadowColor: "#000",
		shadowOpacity: 0.25,
		shadowRadius: 4,
		shadowOffset: { width: 0, height: 2 }
	},
	// "code" chevrons turned upright read as an up/down drag affordance.
	thumbIcon: {
		transform: [{ rotate: "90deg" }]
	},
	bubble: {
		position: "absolute",
		right: THUMB_WIDTH + 16,
		top: 0,
		height: BUBBLE_HEIGHT,
		minWidth: 48,
		paddingHorizontal: 14,
		borderRadius: BUBBLE_HEIGHT / 2,
		alignItems: "center",
		justifyContent: "center",
		elevation: 6
	},
	marker: {
		position: "absolute",
		right: THUMB_WIDTH + 10,
		paddingHorizontal: 6,
		paddingVertical: 1,
		borderRadius: 6,
		opacity: 0.9
	}
})

export default FastScroller
