import { useEffect, useRef, useState, type PointerEvent as ReactPointerEvent, type RefObject } from "react"
import { useTranslation } from "react-i18next"
import type { ClientPoint, DocxEditorHandle } from "@filen/office/react/docx/editor"
import { cn } from "@filen/shared"

// Touch selection over the paged editor, drawn here (the package only reports where the selection's ends are
// and moves them): a teardrop handle under each end to drag it, and a compact menu above the selection with
// Cut, Copy, Paste and Select all, as the platforms' own text selection offers.

// The menu keeps this far from the selection (px), and stays this far inside the editor's box.
const MENU_GAP = 10
const EDGE = 8
const HANDLE_SIZE = 22

interface Placed {
	// The selection it was read for.
	key: string
	start: ClientPoint | undefined
	end: ClientPoint | undefined
	// The editor's box on screen.
	box: { left: number; top: number; width: number; height: number }
}

function samePoint(a: ClientPoint | undefined, b: ClientPoint | undefined): boolean {
	return a === b || (a !== undefined && a.x === b?.x && a.y === b.y && a.height === b.height)
}

function samePlaced(a: Placed | null, b: Placed | null): boolean {
	return (
		a === b ||
		(a !== null &&
			b !== null &&
			samePoint(a.start, b.start) &&
			samePoint(a.end, b.end) &&
			a.box.left === b.box.left &&
			a.box.top === b.box.top &&
			a.box.width === b.box.width &&
			a.box.height === b.box.height)
	)
}

export interface TouchMenuActions {
	cut: (() => void) | null
	copy: (() => void) | null
	paste: (() => void) | null
	selectAll: () => void
}

export function DocxTouchSelection({
	handleRef,
	container,
	shown,
	selectionKey,
	actions
}: {
	handleRef: RefObject<DocxEditorHandle | null>
	// The positioned element the handles and menu are placed in (the editor's own box).
	container: HTMLElement | null
	// The selection was made by touch and holds text.
	shown: boolean
	// Changes whenever the selection moves (its ends may move on screen).
	selectionKey: string
	actions: TouchMenuActions
}) {
	const { t } = useTranslation("preview")
	const [placed, setPlaced] = useState<Placed | null>(null)
	const [dragging, setDragging] = useState<"start" | "end" | null>(null)
	const dragOffset = useRef(0)

	// Where the ends are on screen, read once per frame while shown: after a selection change, a scroll of
	// the pages or a resize.
	useEffect(() => {
		if (!shown || container === null) {
			return
		}

		let frame = 0
		const read = (): void => {
			frame = 0

			const handle = handleRef.current

			if (handle === null) {
				return
			}

			const rect = container.getBoundingClientRect()
			const points = handle.selectionPoints()
			const next: Placed = {
				key: selectionKey,
				start: points.start,
				end: points.end,
				box: { left: rect.left, top: rect.top, width: rect.width, height: rect.height }
			}

			setPlaced(previous => (samePlaced(previous, next) ? previous : next))
		}
		const schedule = (): void => {
			if (frame === 0) {
				frame = requestAnimationFrame(read)
			}
		}
		const scroller = handleRef.current?.scrollElement() ?? null
		const observer = new ResizeObserver(schedule)

		schedule()
		scroller?.addEventListener("scroll", schedule, { passive: true })
		observer.observe(container)

		return () => {
			scroller?.removeEventListener("scroll", schedule)
			observer.disconnect()

			if (frame !== 0) {
				cancelAnimationFrame(frame)
			}
		}
	}, [shown, container, handleRef, selectionKey])

	// Until the ends of a new selection were read, nothing (rather than the last selection's).
	if (!shown || placed?.key !== selectionKey) {
		return null
	}

	const { box } = placed
	const inside = (point: ClientPoint | undefined): point is ClientPoint =>
		point !== undefined &&
		point.y >= box.top &&
		point.y - point.height <= box.top + box.height &&
		point.x >= box.left - 1 &&
		point.x <= box.left + box.width + 1

	const onDown = (end: "start" | "end", point: ClientPoint) => (event: ReactPointerEvent<HTMLDivElement>) => {
		event.preventDefault()
		event.stopPropagation()
		event.currentTarget.setPointerCapture(event.pointerId)
		// The finger holds the drop under the line; the end follows the line's middle above it.
		dragOffset.current = point.y - point.height / 2 - event.clientY
		setDragging(end)
	}
	const onMove = (end: "start" | "end") => (event: ReactPointerEvent<HTMLDivElement>) => {
		if (dragging !== end) {
			return
		}

		handleRef.current?.dragSelectionEnd(end, event.clientX, event.clientY + dragOffset.current)
	}
	const onUp = (event: ReactPointerEvent<HTMLDivElement>) => {
		if (event.currentTarget.hasPointerCapture(event.pointerId)) {
			event.currentTarget.releasePointerCapture(event.pointerId)
		}

		handleRef.current?.endSelectionDrag()
		setDragging(null)
	}

	const handles = (["start", "end"] as const).map(end => {
		const point = placed[end]

		if (!inside(point)) {
			return null
		}

		return (
			// A touch affordance only: assistive technology selects through the editor's own text view.
			<div
				key={end}
				aria-hidden="true"
				className="pointer-events-auto absolute touch-none"
				style={{
					left: point.x - box.left - HANDLE_SIZE / 2,
					top: point.y - box.top,
					width: HANDLE_SIZE,
					height: HANDLE_SIZE + 6
				}}
				onPointerDown={onDown(end, point)}
				onPointerMove={onMove(end)}
				onPointerUp={onUp}
				onPointerCancel={onUp}
			>
				{/* The stem runs up along the line; the drop hangs under it, pointing at the selection. */}
				<span
					aria-hidden="true"
					className="absolute left-1/2 w-0.5 -translate-x-1/2 bg-primary"
					style={{ top: -point.height, height: point.height }}
				/>
				<span
					aria-hidden="true"
					className={cn(
						"absolute top-1 left-1/2 size-3.5 -translate-x-1/2 bg-primary shadow-sm",
						end === "start" ? "rounded-full rounded-tr-none" : "rounded-full rounded-tl-none"
					)}
				/>
			</div>
		)
	})

	// Above the selection's first line, or below its last when there is no room above.
	const anchor = inside(placed.start) ? placed.start : inside(placed.end) ? placed.end : undefined
	const items = [
		{ key: "cut", label: t("previewDocxCut"), run: actions.cut },
		{ key: "copy", label: t("previewDocxCopy"), run: actions.copy },
		{ key: "paste", label: t("previewDocxPaste"), run: actions.paste },
		{ key: "selectAll", label: t("previewDocxSelectAll"), run: actions.selectAll }
	].filter(item => item.run !== null)
	let menu = null

	if (anchor !== undefined && dragging === null) {
		const above = anchor.y - anchor.height - box.top - MENU_GAP
		const below = (placed.end?.y ?? anchor.y) - box.top + HANDLE_SIZE + MENU_GAP
		const top = above >= 44 ? above - 36 : below
		const left = Math.min(Math.max(EDGE, anchor.x - box.left - 80), Math.max(EDGE, box.width - 280))

		menu = (
			<div
				role="menu"
				data-preview-surface
				className="pointer-events-auto absolute flex h-9 items-center overflow-hidden rounded-xl bg-popover text-xs text-popover-foreground shadow-md ring-1 ring-foreground/10"
				style={{ left, top }}
				// Pressing the menu keeps the editor's focus and selection.
				onPointerDown={event => {
					event.preventDefault()
				}}
			>
				{items.map((item, index) => (
					<button
						key={item.key}
						type="button"
						role="menuitem"
						className={cn("h-full px-3 whitespace-nowrap active:bg-accent", index > 0 && "border-l border-border")}
						onClick={() => {
							item.run?.()
						}}
					>
						{item.label}
					</button>
				))}
			</div>
		)
	}

	return (
		<div className="pointer-events-none absolute inset-0 z-10 overflow-hidden">
			{handles}
			{menu}
		</div>
	)
}
