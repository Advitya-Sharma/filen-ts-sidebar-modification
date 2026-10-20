import {
	useEffect,
	useId,
	useImperativeHandle,
	useLayoutEffect,
	useRef,
	useState,
	useSyncExternalStore,
	type FocusEvent,
	type KeyboardEvent,
	type MouseEvent,
	type PointerEvent,
	type Ref
} from "react"
import { flushSync } from "react-dom"
import { isTypedCharacter } from "../../core/sheet/grid/navigation"
import { cellName } from "../../core/sheet/refs/cellRef"
import { cycleReference } from "../../core/sheet/refs/formulaRefs"
import { clampSize, resizeTargets } from "../../core/sheet/sizes/sizes"
import { SHEET_MAX_COLS, SHEET_MAX_ROWS, type SheetAxis, type SheetFormatPatch, type SheetRange } from "../../core/sheet/protocol"
import type { FillDirection } from "../../core/sheet/fill/fill"
import { autofitSizes } from "./autofit"
import { isImeKeydown } from "./ime"
import {
	APP_SHORTCUTS,
	EDITOR_SHORTCUTS,
	SHORTCUT_FORMATS,
	dateEntry,
	isMacPlatform,
	layoutKey,
	shortcutOf,
	timeEntry,
	type SheetShortcut
} from "./keys"
import { canEditSnapshot, type GridResizeDraft, type SheetGridModel, type SheetGridSizes } from "./model"
import { keyMove, stepFrom } from "./navigation"
import { GridRenderer, type GridHit } from "./renderer"
import { addSelection, axisSelection, cellSelection, extendSelection, wholeAxis, type GridPoint, type GridSelection } from "./selection"
import { entryInput, entryText } from "./tiles"

// The grid: canvases drawn by GridRenderer under a scroller that holds only an empty extent, so no cell is a
// React component or a DOM node. React renders what changes with the selection or the entry: the in-cell editor
// (a textarea, so an input method composes in place), the accessible active cell the grid's
// aria-activedescendant names, a live region saying the active cell, and a resize's size tip. Unstyled: the look
// is css/sheet.css and its --fo-sheet-* variables.

export interface SheetGridHandle {
	focus(): void
	// Scrolls a cell into view.
	reveal(point: GridPoint): void
	// Selects a cell and scrolls to it (the name box, Ctrl+G).
	goTo(point: GridPoint): void
	hit(clientX: number, clientY: number): GridHit | null
	// The renderer's frame costs while `measure` is on (perf harness).
	readonly renderer: GridRenderer
}

export interface SheetGridProps {
	readonly grid: SheetGridModel
	readonly label: string
	readonly editorLabel: string
	readonly className?: string
	// What the live region says for the active cell; "<name> <text>" without it.
	readonly describeCell?: (name: string, text: string) => string
	// A resize's size tip; none without it.
	readonly resizeLabel?: (axis: SheetAxis, px: number) => string
	// Sees every grid key first: true takes it (clipboard, shortcuts the app owns).
	readonly onKey?: (event: KeyboardEvent<HTMLDivElement>) => boolean
	// Sees every editor key first: true takes it (formula autocomplete).
	readonly onEditorKey?: (event: KeyboardEvent<HTMLTextAreaElement>) => boolean
	// Sees every keyboard shortcut (keys.ts) first: true takes it. The ones only an app can do (Find, Save, dialogs)
	// do nothing unless it takes them.
	readonly onShortcut?: (shortcut: SheetShortcut) => boolean
	// Command is the primary modifier, as on a Mac (detected from the platform when absent).
	readonly mac?: boolean
	// A double-click (or double-tap) on a cell: true takes it, else the cell's entry opens.
	readonly onCellActivate?: (point: GridPoint) => boolean
	// Change it when the app's theme changes: the canvas reads the --fo-sheet-* variables again.
	readonly themeKey?: string | number
	// The app's size layer, for sessions that cannot hold sizes (text files, view-only workbooks): hand in a new
	// object whenever its sizes change.
	readonly sizes?: SheetGridSizes
	readonly ref?: Ref<SheetGridHandle>
}

interface Drag {
	readonly pointerId: number
	readonly kind: "cells" | "rows" | "cols"
	readonly add: boolean
}

interface Resize {
	readonly pointerId: number
	readonly axis: SheetAxis
	readonly index: number
	readonly targets: readonly number[]
	// Pixels at 100 %.
	readonly start: number
	readonly origin: number
	readonly scale: number
	draft: GridResizeDraft | null
}

interface Tap {
	readonly pointerId: number
	readonly x: number
	readonly y: number
	readonly time: number
}

// How near a header edge a pointer resizes, in pixels.
const EDGE_SLOP = 4
// A touch that moved less than this, quicker than TAP_MS, is a tap.
const TAP_SLOP = 10
const TAP_MS = 500
// Selection drag past the edge scrolls this share of the overshoot per frame.
const AUTOSCROLL_RATE = 0.4

const TEXT_KINDS: ReadonlySet<string> = new Set(["csv", "tsv", "tab"])

const FORMAT_SHORTCUTS: ReadonlySet<SheetShortcut> = new Set<SheetShortcut>([
	"bold",
	"italic",
	"underline",
	"strikethrough",
	"formatGeneral",
	"formatCurrency",
	"formatPercent",
	"formatScientific",
	"formatDate",
	"formatTime",
	"formatNumber"
])

// Cells a move on a protected sheet passes over looking for one it may select.
const MAX_PROTECTED_STEPS = 1024

// A command whose answer the grid needs nothing from: a refusal reaches onRefusal, a session that died its
// "failed" event.
function settled(promise: Promise<unknown>): void {
	promise.then(undefined, () => undefined)
}

function defaultDescription(name: string, text: string): string {
	return text === "" ? name : `${name} ${text}`
}

interface CellEditorProps {
	readonly id: string
	readonly label: string
	readonly text: string
	readonly invalid: boolean
	readonly element: (element: HTMLTextAreaElement | null) => void
	readonly onText: (text: string) => void
	readonly onKeyDown: (event: KeyboardEvent<HTMLTextAreaElement>) => void
	readonly onBlur: (event: FocusEvent<HTMLTextAreaElement>) => void
}

// The in-cell editor. It exists only while an entry is typed, and takes focus as it appears (in the same task, so
// a composition started by the key that opened it lands here), with the caret after what it starts from.
function CellEditor({ id, label, text, invalid, element, onText, onKeyDown, onBlur }: CellEditorProps) {
	const inputRef = useRef<HTMLTextAreaElement>(null)

	useLayoutEffect(() => {
		const input = inputRef.current

		input?.focus({ preventScroll: true })
		input?.setSelectionRange(input.value.length, input.value.length)
	}, [])

	useLayoutEffect(() => {
		element(inputRef.current)

		return () => element(null)
	}, [element])

	return (
		<textarea
			ref={inputRef}
			id={id}
			data-sheet-editor=""
			className="fo-sheet-editor"
			aria-label={label}
			aria-invalid={invalid}
			spellCheck={false}
			autoComplete="off"
			rows={1}
			value={text}
			onChange={event => onText(event.target.value)}
			onKeyDown={onKeyDown}
			onBlur={onBlur}
		/>
	)
}

export function SheetGrid({
	grid,
	label,
	editorLabel,
	className,
	describeCell,
	resizeLabel,
	onKey,
	onEditorKey,
	onShortcut,
	mac,
	onCellActivate,
	themeKey,
	sizes,
	ref
}: SheetGridProps) {
	const snapshot = useSyncExternalStore(grid.subscribe, grid.getSnapshot, grid.getSnapshot)
	const [owned, setOwned] = useState(() => ({ grid, renderer: new GridRenderer(grid) }))
	let current = owned

	// A new model replaces the renderer during render; the old one detaches with its effect.
	if (owned.grid !== grid) {
		current = { grid, renderer: new GridRenderer(grid) }
		setOwned(current)
	}

	const renderer = current.renderer
	const rootRef = useRef<HTMLDivElement>(null)
	const scrollerRef = useRef<HTMLDivElement>(null)
	const extentRef = useRef<HTMLDivElement>(null)
	const cellsRef = useRef<HTMLCanvasElement>(null)
	const chromeRef = useRef<HTMLCanvasElement>(null)
	const activeRef = useRef<HTMLDivElement>(null)
	const tipRef = useRef<HTMLDivElement>(null)
	const dragRef = useRef<Drag | null>(null)
	const resizeRef = useRef<Resize | null>(null)
	const tapRef = useRef<Tap | null>(null)
	const autoscrollRef = useRef<{ frame: number; x: number; y: number } | null>(null)
	const ids = useId()
	const activeId = `${ids}-active`
	const editorId = `${ids}-editor`
	const { selection, editing, active, layout } = snapshot
	const describe = describeCell ?? defaultDescription
	const liveText = describe(cellName(selection.active.row, selection.active.col), active.text)
	const readOnly = !canEditSnapshot(snapshot)
	const onMac = mac ?? isMacPlatform()

	useLayoutEffect(() => {
		const root = rootRef.current
		const scroller = scrollerRef.current
		const extent = extentRef.current
		const cells = cellsRef.current
		const chrome = chromeRef.current
		const activeCell = activeRef.current

		if (root === null || scroller === null || extent === null || cells === null || chrome === null || activeCell === null)
			return undefined

		renderer.attach({ root, scroller, extent, cells, chrome, active: activeCell })

		// Rows fit the wrapped text an edit changes, measured as the renderer draws it (rowFit.ts).
		const unmeasure = grid.measureWith({ measurer: () => renderer.measurerFor(), font: () => renderer.themeFont() })

		return () => {
			unmeasure()
			renderer.dispose()
		}
	}, [grid, renderer])

	useEffect(() => {
		if (themeKey !== undefined) renderer.refreshTheme()
	}, [renderer, themeKey])

	// Before the frame, so a resize's draft never shows the old size for a frame.
	useLayoutEffect(() => {
		grid.setSizes(sizes ?? null)
	}, [grid, sizes])

	useEffect(
		() => () => {
			const autoscroll = autoscrollRef.current

			if (autoscroll !== null) cancelAnimationFrame(autoscroll.frame)
		},
		[]
	)

	useImperativeHandle(
		ref,
		() => ({
			focus: () => scrollerRef.current?.focus({ preventScroll: true }),
			reveal: point => renderer.reveal(point),
			goTo: point => {
				const current = grid.getSnapshot().layout

				if (grid.choose(cellSelection(point, current === null ? [] : grid.merges(current)))) renderer.reveal(point)
			},
			hit: (clientX, clientY) => renderer.hit(clientX, clientY, EDGE_SLOP),
			renderer
		}),
		[grid, renderer]
	)

	function merges(): readonly SheetRange[] {
		const current = grid.getSnapshot().layout

		return current === null ? [] : grid.merges(current)
	}

	function focusGrid(): void {
		scrollerRef.current?.focus({ preventScroll: true })
	}

	// A selection the user made: none on a cell sheet protection keeps from being selected.
	function select(next: GridSelection, reveal: GridPoint | null): void {
		if (grid.choose(next) && reveal !== null) renderer.reveal(reveal)
	}

	// The active cell for a click on a column (row) header: the first row (column) on screen.
	function headerActive(kind: "rows" | "cols", index: number): GridPoint {
		const view = renderer.currentView()

		return kind === "cols"
			? { row: view === null ? 0 : view.rows.frozen > 0 ? view.rows.frozenFirst : view.rows.first, col: index }
			: { row: index, col: view === null ? 0 : view.cols.frozen > 0 ? view.cols.frozenFirst : view.cols.first }
	}

	function selectHit(hit: GridHit, shift: boolean, add: boolean): Drag["kind"] | null {
		const current = grid.getSnapshot().selection
		// Shift extends the last range and keeps the others.
		const earlier = current.ranges.length > 1 ? { ...current, ranges: current.ranges.slice(0, -1) } : null

		switch (hit.kind) {
			case "corner":
				select(
					{
						active: { row: 0, col: 0 },
						anchor: { row: 0, col: 0 },
						focus: { row: SHEET_MAX_ROWS - 1, col: SHEET_MAX_COLS - 1 },
						ranges: [{ startRow: 0, startCol: 0, endRow: SHEET_MAX_ROWS - 1, endCol: SHEET_MAX_COLS - 1 }]
					},
					null
				)

				return null
			case "colHeader":
				if (shift) select(axisSelection("cols", current.anchor.col, hit.col, current.active, earlier), null)
				else select(axisSelection("cols", hit.col, hit.col, headerActive("cols", hit.col), add ? current : null), null)

				return "cols"
			case "rowHeader":
				if (shift) select(axisSelection("rows", current.anchor.row, hit.row, current.active, earlier), null)
				else select(axisSelection("rows", hit.row, hit.row, headerActive("rows", hit.row), add ? current : null), null)

				return "rows"
			case "cell": {
				const point = { row: hit.row, col: hit.col }

				if (shift) select(extendSelection(current, point, merges()), null)
				else if (add) select(addSelection(current, point, merges()), null)
				else select(cellSelection(point, merges()), null)

				return "cells"
			}
		}
	}

	// --- Resizing ---

	function canResize(): boolean {
		return grid.canResize()
	}

	function targetsFor(axis: SheetAxis, index: number): number[] {
		const { selection: current, sheet } = grid.getSnapshot()
		const range = current.ranges[current.ranges.length - 1] as SheetRange
		const axes = grid.axes(sheet)
		const whole =
			axis === "cols"
				? range.startRow === 0 && range.endRow >= SHEET_MAX_ROWS - 1
				: range.startCol === 0 && range.endCol >= SHEET_MAX_COLS - 1

		return resizeTargets(axis, index, range, whole, at =>
			axes === null ? false : (axis === "cols" ? axes.cols : axes.rows).hidden(at)
		)
	}

	function startResize(event: PointerEvent<HTMLDivElement>, axis: SheetAxis, index: number): void {
		const { sheet } = grid.getSnapshot()
		const axes = grid.axes(sheet)

		if (axes === null) return

		event.preventDefault()
		event.currentTarget.setPointerCapture(event.pointerId)
		resizeRef.current = {
			pointerId: event.pointerId,
			axis,
			index,
			targets: targetsFor(axis, index),
			start: (axis === "cols" ? axes.cols : axes.rows).size(index) / axes.scale,
			origin: axis === "cols" ? event.clientX : event.clientY,
			scale: axes.scale,
			draft: null
		}
		renderer.setGuide({ axis, index })
	}

	// A column's edge in a right-to-left sheet is its left one: dragging it left widens the column.
	function resizeSize(resize: Resize, clientX: number, clientY: number): number {
		const moved = (resize.axis === "cols" ? clientX : clientY) - resize.origin
		const mirrored = resize.axis === "cols" && grid.layout(grid.getSnapshot().sheet)?.rightToLeft === true

		return clampSize(resize.axis, resize.start + (mirrored ? -moved : moved) / resize.scale)
	}

	function moveResize(resize: Resize, event: PointerEvent<HTMLDivElement>): void {
		const size = resizeSize(resize, event.clientX, event.clientY)
		const draft: GridResizeDraft = {
			sheet: grid.getSnapshot().sheet,
			axis: resize.axis,
			sizes: new Map(resize.targets.map(target => [target, size]))
		}

		resize.draft = draft
		grid.setDraft(draft)

		const tip = tipRef.current
		const local = renderer.local(event.clientX, event.clientY)

		if (tip !== null && local !== null && resizeLabel !== undefined) {
			tip.textContent = resizeLabel(resize.axis, size)
			tip.hidden = false
			tip.style.transform = `translate(${String(local.x + 12)}px, ${String(local.y + 12)}px)`
		}
	}

	function endResize(commit: PointerEvent<HTMLDivElement> | null): void {
		const resize = resizeRef.current

		resizeRef.current = null
		renderer.setGuide(null)

		if (tipRef.current !== null) tipRef.current.hidden = true
		if (resize === null) return

		const size = commit === null ? resize.start : resizeSize(resize, commit.clientX, commit.clientY)

		// A press that never moved (a click, or the first half of a double-click) resizes nothing.
		if (commit === null || Math.abs(size - Math.round(resize.start)) < 1) {
			grid.setDraft(null)

			return
		}

		const draft = resize.draft
		const sheet = grid.getSnapshot().sheet

		// The draft stays until the new sizes are the sheet's (the session's or the app's), so nothing flickers back.
		const done = (): void => {
			if (grid.getDraft() === draft) grid.setDraft(null)
		}

		grid.resize(sheet, resize.axis, Uint32Array.from(resize.targets), new Float32Array(resize.targets.length).fill(size)).then(
			done,
			done
		)
	}

	function autofit(axis: SheetAxis, index: number): void {
		const measurer = renderer.measurerFor()

		if (measurer === null || !canResize()) return

		const sheet = grid.getSnapshot().sheet

		void autofitSizes(grid, sheet, axis, targetsFor(axis, index), measurer, renderer.themeFont()).then(
			({ indices, sizes }) => {
				if (indices.length > 0) settled(grid.resize(sheet, axis, indices, sizes))
			},
			() => undefined
		)
	}

	// --- Pointer ---

	function stopAutoscroll(): void {
		const autoscroll = autoscrollRef.current

		if (autoscroll !== null) cancelAnimationFrame(autoscroll.frame)

		autoscrollRef.current = null
	}

	function extendDrag(drag: Drag, row: number | null, col: number | null): void {
		const current = grid.getSnapshot().selection
		const focus = { row: row ?? current.focus.row, col: col ?? current.focus.col }

		if (drag.kind === "cells") {
			if (focus.row !== current.focus.row || focus.col !== current.focus.col) grid.select(extendSelection(current, focus, merges()))
		} else {
			const kind = drag.kind
			const index = kind === "cols" ? focus.col : focus.row
			const start = kind === "cols" ? current.anchor.col : current.anchor.row
			const previous = drag.add ? { ...current, ranges: current.ranges.slice(0, -1) } : null

			grid.select(axisSelection(kind, start, index, current.active, previous))
		}
	}

	// While a selection drag is past the grid's edge, scroll towards it and extend to the cell at the edge.
	function autoscroll(x: number, y: number): void {
		const view = renderer.currentView()

		if (view === null) return

		const left = view.cols.header + view.cols.frozenSize
		const top = view.rows.header + view.rows.frozenSize
		const dx = x < left ? x - left : x > view.width ? x - view.width : 0
		const dy = y < top ? y - top : y > view.height ? y - view.height : 0

		if (dx === 0 && dy === 0) {
			stopAutoscroll()

			return
		}

		if (autoscrollRef.current !== null) {
			autoscrollRef.current.x = x
			autoscrollRef.current.y = y

			return
		}

		const tick = (): void => {
			const state = autoscrollRef.current
			const drag = dragRef.current
			const now = renderer.currentView()

			if (state === null || drag === null || now === null) return

			const ddx = state.x < left ? state.x - left : state.x > now.width ? state.x - now.width : 0
			const ddy = state.y < top ? state.y - top : state.y > now.height ? state.y - now.height : 0

			renderer.scrollBy(drag.kind === "rows" ? 0 : ddx * AUTOSCROLL_RATE, drag.kind === "cols" ? 0 : ddy * AUTOSCROLL_RATE)

			const edge = renderer.currentView()

			if (edge !== null) {
				const row = ddy < 0 ? edge.rows.first : ddy > 0 ? edge.rows.last : null
				const col = ddx < 0 ? edge.cols.first : ddx > 0 ? edge.cols.last : null

				extendDrag(drag, row, col)
			}

			state.frame = requestAnimationFrame(tick)
		}

		autoscrollRef.current = { frame: requestAnimationFrame(tick), x, y }
	}

	function handlePointerDown(event: PointerEvent<HTMLDivElement>): void {
		if (event.button !== 0 && event.button !== 2) return

		const hit = renderer.hit(event.clientX, event.clientY, EDGE_SLOP)

		if (hit === null) return

		if (event.pointerType === "touch") {
			// Touch pans natively: a tap selects on release.
			tapRef.current = { pointerId: event.pointerId, x: event.clientX, y: event.clientY, time: event.timeStamp }

			return
		}

		// A right-click acts on the selection it lands in, or on the cell it lands on.
		if (event.button === 2) {
			if (hit.kind === "cell") {
				const current = grid.getSnapshot().selection
				const inside = current.ranges.some(
					range => hit.row >= range.startRow && hit.row <= range.endRow && hit.col >= range.startCol && hit.col <= range.endCol
				)

				if (!inside) selectHit(hit, false, false)
			}

			return
		}

		focusGrid()

		if ((hit.kind === "colHeader" || hit.kind === "rowHeader") && hit.edge !== null && canResize()) {
			startResize(event, hit.kind === "colHeader" ? "cols" : "rows", hit.edge)

			return
		}

		const add = event.ctrlKey || event.metaKey
		const kind = selectHit(hit, event.shiftKey, add)

		if (kind !== null) {
			dragRef.current = { pointerId: event.pointerId, kind, add }
			event.currentTarget.setPointerCapture(event.pointerId)
		}
	}

	function handlePointerMove(event: PointerEvent<HTMLDivElement>): void {
		const resize = resizeRef.current

		if (resize !== null && resize.pointerId === event.pointerId) {
			moveResize(resize, event)

			return
		}

		const drag = dragRef.current

		if (drag === null || drag.pointerId !== event.pointerId) {
			if (event.pointerType !== "touch") hoverCursor(event)

			return
		}

		const local = renderer.local(event.clientX, event.clientY)

		if (local !== null) autoscroll(renderer.layoutX(local.x), local.y)

		const hit = renderer.hit(event.clientX, event.clientY, 0)

		if (hit === null) return

		if (hit.kind === "cell") extendDrag(drag, drag.kind === "cols" ? null : hit.row, drag.kind === "rows" ? null : hit.col)
		else if (hit.kind === "colHeader" && drag.kind !== "rows") extendDrag(drag, null, hit.col)
		else if (hit.kind === "rowHeader" && drag.kind !== "cols") extendDrag(drag, hit.row, null)
	}

	function hoverCursor(event: PointerEvent<HTMLDivElement>): void {
		const root = rootRef.current
		const hit = renderer.hit(event.clientX, event.clientY, EDGE_SLOP)
		const cursor =
			hit?.kind === "colHeader" && hit.edge !== null && canResize()
				? "col-resize"
				: hit?.kind === "rowHeader" && hit.edge !== null && canResize()
					? "row-resize"
					: ""

		if (root !== null && root.dataset["cursor"] !== cursor) root.dataset["cursor"] = cursor
	}

	function handlePointerUp(event: PointerEvent<HTMLDivElement>): void {
		const resize = resizeRef.current

		if (resize !== null && resize.pointerId === event.pointerId) {
			endResize(event)

			return
		}

		const tap = tapRef.current

		if (tap !== null && tap.pointerId === event.pointerId) {
			tapRef.current = null

			if (Math.hypot(event.clientX - tap.x, event.clientY - tap.y) < TAP_SLOP && event.timeStamp - tap.time < TAP_MS) {
				const hit = renderer.hit(event.clientX, event.clientY, 0)

				if (hit !== null) {
					focusGrid()
					selectHit(hit, false, false)
				}
			}

			return
		}

		if (dragRef.current?.pointerId === event.pointerId) {
			dragRef.current = null
			stopAutoscroll()
		}
	}

	function handlePointerCancel(event: PointerEvent<HTMLDivElement>): void {
		if (resizeRef.current?.pointerId === event.pointerId) endResize(null)
		if (tapRef.current?.pointerId === event.pointerId) tapRef.current = null
		if (dragRef.current?.pointerId === event.pointerId) {
			dragRef.current = null
			stopAutoscroll()
		}
	}

	function handleDoubleClick(event: MouseEvent<HTMLDivElement>): void {
		const hit = renderer.hit(event.clientX, event.clientY, EDGE_SLOP)

		if (hit === null) return

		if ((hit.kind === "colHeader" || hit.kind === "rowHeader") && hit.edge !== null) {
			autofit(hit.kind === "colHeader" ? "cols" : "rows", hit.edge)

			return
		}

		if (hit.kind !== "cell") return

		const point = { row: hit.row, col: hit.col }

		if (onCellActivate?.(point) === true) return

		grid.startEditing(null, "edit")
	}

	// --- Keys ---

	function navigate(direction: "up" | "down" | "left" | "right", extend: boolean): void {
		const { sheet, selection: current } = grid.getSnapshot()
		const from = extend ? current.focus : current.active

		grid.session.query("navigate", { sheet, row: from.row, col: from.col, direction, to: "dataEdge" }).then(
			to => {
				const now = grid.getSnapshot()

				if (now.sheet !== sheet || now.selection !== current) return

				const point = { row: to.row, col: to.col }

				select(extend ? extendSelection(current, point, merges()) : cellSelection(point, merges()), point)
			},
			() => undefined
		)
	}

	function clearContents(): void {
		const { sheet, selection: current } = grid.getSnapshot()

		if (!grid.canEdit()) return

		void grid.intent()
		settled(grid.command({ type: "cells.clear", sheet, ranges: current.ranges, what: "contents" }))
	}

	// Sends an edit, once the session was told the user set out to edit.
	function edit(run: () => Promise<unknown> | null): void {
		if (!grid.canChange()) return

		void grid.intent()

		const sent = run()

		if (sent !== null) settled(sent)
	}

	// On or off as the active cell is, as a toolbar button toggles it.
	function toggleFormat(format: "bold" | "italic" | "underline" | "strike"): void {
		const { sheet, selection: current } = grid.getSnapshot()
		const { row, col } = current.active
		const apply = (style: number): void => {
			const look = grid.style(style)
			const on = format === "underline" ? look.underline !== undefined : look[format] === true
			const patch: SheetFormatPatch = format === "underline" ? { underline: on ? null : "single" } : { [format]: !on }

			edit(() => grid.command({ type: "format.apply", sheet, ranges: current.ranges, patch }))
		}
		const held = grid.entry(sheet, row, col)

		if (held !== null || grid.tileHeld(sheet, row, col)) apply(held === null ? 0 : (held.tile.styles[held.entry] ?? 0))
		else
			grid.session.query("cell", { sheet, row, col }).then(
				detail => apply(detail.style),
				() => undefined
			)
	}

	// Ctrl+A: the data region round the active cell, then (pressed again, or with no data round it) the whole sheet.
	function selectRegion(whole: boolean): void {
		const { sheet, selection: current } = grid.getSnapshot()
		const all: GridSelection = {
			active: current.active,
			anchor: { row: 0, col: 0 },
			focus: { row: SHEET_MAX_ROWS - 1, col: SHEET_MAX_COLS - 1 },
			ranges: [{ startRow: 0, startCol: 0, endRow: SHEET_MAX_ROWS - 1, endCol: SHEET_MAX_COLS - 1 }]
		}

		grid.session.query("dataRegion", { sheet, row: current.active.row, col: current.active.col }).then(
			region => {
				const now = grid.getSnapshot()

				if (now.sheet !== sheet || now.selection !== current) return

				const last = current.ranges[current.ranges.length - 1]
				const alone =
					region.startRow === region.endRow &&
					region.startCol === region.endCol &&
					grid.entry(sheet, region.startRow, region.startCol) === null
				const same =
					last !== undefined &&
					current.ranges.length === 1 &&
					last.startRow === region.startRow &&
					last.startCol === region.startCol &&
					last.endRow === region.endRow &&
					last.endCol === region.endCol

				if (whole && (alone || same)) {
					select(all, null)

					return
				}

				select(
					{
						active: current.active,
						anchor: { row: region.startRow, col: region.startCol },
						focus: { row: region.endRow, col: region.endCol },
						ranges: [region]
					},
					null
				)
			},
			() => undefined
		)
	}

	// Ctrl+' and Ctrl+Shift+": an entry starts from the cell above's formula (as typed there) or its value.
	function fromAbove(value: boolean): void {
		const { sheet, selection: current } = grid.getSnapshot()
		const { row, col } = current.active

		if (row === 0 || !grid.canEdit()) return

		const held = grid.entry(sheet, row - 1, col)

		if (held !== null || grid.tileHeld(sheet, row - 1, col)) {
			const text = held === null ? "" : value ? entryText(held.tile, held.entry) : entryInput(held.tile, held.entry)

			grid.startEditing(text, "edit")

			return
		}

		grid.session.query("cell", { sheet, row: row - 1, col }).then(
			detail => {
				if (grid.getSnapshot().selection === current) grid.startEditing(detail.input, "edit")
			},
			() => undefined
		)
	}

	function switchSheet(step: 1 | -1): void {
		const { sheets, sheet } = grid.getSnapshot()
		const shown = sheets.filter(entry => entry.state === "visible")
		const target = shown[shown.findIndex(entry => entry.id === sheet) + step]

		if (target === undefined) return

		void grid.commitEditing(null).then(done => {
			if (done) grid.setSheet(target.id)
		})
	}

	// Rows or columns inserted or deleted when whole ones are selected; cells alone are the app's (a dialog asks how,
	// as the app was handed the shortcut first). The key is taken either way: the page never zooms on it.
	function insertOrDelete(type: "axis.insert" | "axis.delete"): void {
		const { sheet, selection: current } = grid.getSnapshot()
		const range = current.ranges[0]
		const axis = wholeAxis(current)

		if (range === undefined || axis === null) return

		edit(() =>
			grid.command(
				axis === "rows"
					? { type, sheet, axis: "rows", at: range.startRow, count: range.endRow - range.startRow + 1 }
					: { type, sheet, axis: "cols", at: range.startCol, count: range.endCol - range.startCol + 1 }
			)
		)
	}

	// A shortcut the grid performs itself: false for the app's.
	function runShortcut(shortcut: SheetShortcut): boolean {
		const { sheet, selection: current, state } = grid.getSnapshot()
		const merges_ = merges()
		const range = current.ranges[current.ranges.length - 1] as SheetRange
		const earlier = current.ranges.length > 1 ? { ...current, ranges: current.ranges.slice(0, -1) } : null
		const numFmt = SHORTCUT_FORMATS[shortcut]
		// A text file holds no formats.
		const formats = !TEXT_KINDS.has(grid.kind)

		if (FORMAT_SHORTCUTS.has(shortcut) && !formats) return true

		if (numFmt !== undefined) {
			edit(() => grid.command({ type: "format.apply", sheet, ranges: current.ranges, patch: { numFmt } }))

			return true
		}

		switch (shortcut) {
			case "fillDown":
			case "fillRight": {
				const direction: FillDirection = shortcut === "fillDown" ? "down" : "right"

				edit(() => grid.command({ type: "fill.copy", sheet, ranges: current.ranges, direction }))

				return true
			}
			case "repeat":
				edit(() => grid.repeat(current))

				return true
			case "redoOrRepeat":
				edit(() => (state.canRedo ? grid.redo() : grid.repeat(current)))

				return true
			case "insertDate":
			case "insertTime":
				grid.startEditing(shortcut === "insertDate" ? dateEntry(new Date()) : timeEntry(new Date()), "enter")

				return true
			case "selectColumns":
				select(axisSelection("cols", range.startCol, range.endCol, current.active, earlier), null)

				return true
			case "selectRows":
				select(axisSelection("rows", range.startRow, range.endRow, current.active, earlier), null)

				return true
			case "selectAll":
			case "selectRegion":
				selectRegion(shortcut === "selectAll")

				return true
			case "collapseSelection":
				select(cellSelection(current.active, merges_), current.active)

				return true
			case "nextCorner": {
				const corners: GridPoint[] = [
					{ row: range.startRow, col: range.startCol },
					{ row: range.startRow, col: range.endCol },
					{ row: range.endRow, col: range.endCol },
					{ row: range.endRow, col: range.startCol }
				]
				const at = corners.findIndex(corner => corner.row === current.active.row && corner.col === current.active.col)
				const next = corners[(at + 1) % corners.length] as GridPoint

				select({ ...current, active: next }, next)

				return true
			}
			case "revealActive":
				renderer.reveal(current.active)

				return true
			case "screenLeft":
			case "screenRight": {
				const step = renderer.pageCols() * (shortcut === "screenLeft" ? -1 : 1)
				const point = { row: current.active.row, col: Math.max(0, Math.min(SHEET_MAX_COLS - 1, current.active.col + step)) }

				select(cellSelection(point, merges_), point)

				return true
			}
			case "nextSheet":
			case "previousSheet":
				switchSheet(shortcut === "nextSheet" ? 1 : -1)

				return true
			case "formulaFromAbove":
			case "valueFromAbove":
				fromAbove(shortcut === "valueFromAbove")

				return true
			case "showFormulas":
				grid.setShowFormulas(!grid.getSnapshot().showFormulas)

				return true
			case "insertCells":
			case "deleteCells":
				insertOrDelete(shortcut === "insertCells" ? "axis.insert" : "axis.delete")

				return true
			case "bold":
			case "italic":
			case "underline":
				toggleFormat(shortcut)

				return true
			case "strikethrough":
				toggleFormat("strike")

				return true
			default:
				return false
		}
	}

	// Ctrl+Home on a sheet with frozen panes goes to the first cell under and right of them, as Excel does.
	function frozenHome(event: KeyboardEvent<HTMLDivElement>): boolean {
		const pane = grid.getSnapshot().layout?.pane ?? null

		if (event.key !== "Home" || !(event.ctrlKey || event.metaKey) || event.altKey || pane === null || pane.rows + pane.cols === 0)
			return false

		const current = grid.getSnapshot().selection
		const point = { row: pane.rows > 0 ? pane.firstRow + pane.rows : 0, col: pane.cols > 0 ? pane.firstCol + pane.cols : 0 }

		select(event.shiftKey ? extendSelection(current, point, merges()) : cellSelection(point, merges()), point)

		return true
	}

	function handleKeyDown(event: KeyboardEvent<HTMLDivElement>): void {
		if (event.target !== event.currentTarget || onKey?.(event) === true) return

		if (resizeRef.current !== null) {
			if (event.key === "Escape") {
				event.preventDefault()
				event.stopPropagation()
				endResize(null)
			}

			return
		}

		const shortcut = shortcutOf(event, onMac)

		if (shortcut !== null) {
			if (onShortcut?.(shortcut) === true || (!APP_SHORTCUTS.has(shortcut) && runShortcut(shortcut))) {
				event.preventDefault()

				return
			}

			// An app shortcut nobody took: the key keeps its default.
			if (APP_SHORTCUTS.has(shortcut)) return
		}

		if (frozenHome(event)) {
			event.preventDefault()

			return
		}

		const mod = event.ctrlKey || event.metaKey
		const key = layoutKey(event).toLowerCase()

		if (mod && !event.altKey && (key === "z" || key === "y")) {
			if (grid.canChange()) {
				event.preventDefault()

				settled(key === "z" && !event.shiftKey ? grid.undo() : grid.redo())
			}

			return
		}

		const { sheet, layout: current, selection: currentSelection } = grid.getSnapshot()
		const axes = grid.axes(sheet)

		if (current !== null && axes !== null) {
			const move = keyMove(event, currentSelection, axes, current, grid.merges(current), renderer.pageRows())

			if (move !== null) {
				event.preventDefault()

				if (move.kind === "navigate") {
					navigate(move.direction, move.extend)

					return
				}

				// On a protected sheet a move passes over the cells it may not select, as Tab goes from one unlocked cell
				// to the next in Excel; with none that way, the selection stays.
				let landed = move

				for (let step = 0; step < MAX_PROTECTED_STEPS && !grid.selectable(sheet, landed.selection.active); step++) {
					const next = keyMove(event, landed.selection, axes, current, grid.merges(current), renderer.pageRows())

					if (
						next?.kind !== "select" ||
						(next.selection.active.row === landed.selection.active.row &&
							next.selection.active.col === landed.selection.active.col)
					)
						break

					landed = next
				}

				select(landed.selection, landed.reveal)

				return
			}
		}

		if (event.key === "F2") {
			event.preventDefault()
			grid.startEditing(null, "edit")

			return
		}

		if (event.key === "Delete") {
			event.preventDefault()
			clearContents()

			return
		}

		// Backspace empties the active cell into a new entry, as Excel does.
		if (event.key === "Backspace") {
			event.preventDefault()
			grid.startEditing("", "enter")

			return
		}

		// An input method's key, or a dead key: an empty entry opens and takes focus before the key's default runs,
		// so what it composes lands in the entry.
		if (isImeKeydown(event.nativeEvent) || event.key === "Dead") {
			flushSync(() => {
				grid.startEditing("", "enter")
			})

			return
		}

		// A character starts an entry that replaces the cell.
		if (isTypedCharacter(event) && grid.canEdit()) {
			event.preventDefault()
			grid.startEditing(event.key, "enter")
		}
	}

	// --- Editor ---

	function commitTo(next: GridSelection | null): void {
		void grid.commitEditing(next)
		focusGrid()

		if (next !== null) renderer.reveal(next.active)
	}

	function handleEditorKey(event: KeyboardEvent<HTMLTextAreaElement>): void {
		if (onEditorKey?.(event) === true || isImeKeydown(event.nativeEvent)) return

		const { sheet, layout: current, selection: currentSelection, editing: entry } = grid.getSnapshot()
		const axes = grid.axes(sheet)

		if (entry === null || current === null || axes === null) return

		const mergeList = grid.merges(current)
		const shortcut = shortcutOf(event, onMac)

		if (shortcut !== null && EDITOR_SHORTCUTS.has(shortcut)) {
			// Today's date or the time, typed at the caret.
			event.preventDefault()

			const input = event.currentTarget

			input.setRangeText(
				shortcut === "insertDate" ? dateEntry(new Date()) : timeEntry(new Date()),
				input.selectionStart,
				input.selectionEnd,
				"end"
			)
			grid.setEditText(input.value)

			return
		}

		// F4 in a formula: the reference at the caret cycles through absolute and relative.
		if (shortcut === "repeat" && entry.text.startsWith("=")) {
			const input = event.currentTarget
			const cycled = cycleReference(entry.text.slice(1), (input.selectionEnd ?? input.value.length) - 1)

			if (cycled !== null) {
				event.preventDefault()
				input.value = `=${cycled.formula}`
				input.setSelectionRange(cycled.end + 1, cycled.end + 1)
				grid.setEditText(input.value)
			}

			return
		}

		if (event.key === "Enter" && (event.ctrlKey || event.metaKey) && !event.altKey && !event.shiftKey) {
			event.preventDefault()
			fillSelection(entry.text, currentSelection)

			return
		}

		if (event.key === "Enter" && event.altKey) {
			// A line break in the entry.
			event.preventDefault()

			const input = event.currentTarget

			input.setRangeText("\n", input.selectionStart, input.selectionEnd, "end")
			grid.setEditText(input.value)

			return
		}

		if (event.key === "Enter" || event.key === "Tab") {
			event.preventDefault()

			const move = keyMove(
				{ key: event.key, shiftKey: event.shiftKey, ctrlKey: false, metaKey: false, altKey: false },
				currentSelection,
				axes,
				current,
				mergeList,
				1
			)

			commitTo(move?.kind === "select" ? move.selection : null)

			return
		}

		if (event.key === "Escape") {
			event.preventDefault()
			event.stopPropagation()
			grid.cancelEditing()
			focusGrid()

			return
		}

		if (event.key === "F2") {
			event.preventDefault()
			grid.setEditMode(entry.mode === "enter" ? "edit" : "enter")

			return
		}

		const direction = { ArrowUp: "up", ArrowDown: "down", ArrowLeft: "left", ArrowRight: "right" }[event.key] as
			"up" | "down" | "left" | "right" | undefined

		// Enter mode: an arrow commits and moves, as Excel does.
		if (direction !== undefined && entry.mode === "enter" && !event.shiftKey && !event.ctrlKey && !event.metaKey && !event.altKey) {
			event.preventDefault()
			commitTo(stepFrom({ row: entry.row, col: entry.col }, direction, axes, current, mergeList))
		}
	}

	// Ctrl+Enter: the entry goes into every selected cell (a merge's first one), a formula's relative references
	// moving from the active cell to each, as Excel fills a selection. The session lists the cells (fill.entry).
	function fillSelection(text: string, current: GridSelection): void {
		const { sheet } = grid.getSnapshot()

		grid.cancelEditing()
		focusGrid()
		edit(() => grid.command({ type: "fill.entry", sheet, ranges: current.ranges, entry: text, from: current.active }))
	}

	function handleEditorBlur(event: FocusEvent<HTMLTextAreaElement>): void {
		const related = event.relatedTarget

		// Focus went to a control that edits the same entry (a formula bar marks itself so).
		if (related instanceof Element && related.closest("[data-sheet-editor]") !== null) return
		// The window or tab losing focus leaves the entry open: focus comes back to it.
		if (!document.hasFocus()) return
		if (grid.getSnapshot().editing !== null) void grid.commitEditing(null)
	}

	function editorElement(element: HTMLTextAreaElement | null): void {
		renderer.setEditor(element)
	}

	return (
		<div
			ref={rootRef}
			className={className === undefined ? "fo-sheet" : `fo-sheet ${className}`}
			data-editing={editing === null ? undefined : ""}
		>
			<canvas
				ref={cellsRef}
				className="fo-sheet-canvas"
				aria-hidden="true"
			/>
			<canvas
				ref={chromeRef}
				className="fo-sheet-canvas"
				aria-hidden="true"
			/>
			<div
				ref={scrollerRef}
				className="fo-sheet-scroller"
				role="grid"
				aria-label={label}
				aria-rowcount={SHEET_MAX_ROWS}
				aria-colcount={SHEET_MAX_COLS}
				aria-multiselectable="true"
				aria-readonly={readOnly}
				aria-activedescendant={layout === null ? undefined : activeId}
				tabIndex={0}
				onPointerDown={handlePointerDown}
				onPointerMove={handlePointerMove}
				onPointerUp={handlePointerUp}
				onPointerCancel={handlePointerCancel}
				onDoubleClick={handleDoubleClick}
				onKeyDown={handleKeyDown}
			>
				<div
					ref={extentRef}
					className="fo-sheet-extent"
				>
					<div
						ref={activeRef}
						role="row"
						aria-rowindex={selection.active.row + 1}
						className="fo-sheet-a11y"
					>
						<div
							id={activeId}
							role="gridcell"
							aria-colindex={selection.active.col + 1}
							aria-selected="true"
							aria-readonly={readOnly}
						>
							{active.text}
						</div>
					</div>
				</div>
			</div>
			{editing === null ? null : (
				<CellEditor
					id={editorId}
					label={editorLabel}
					text={editing.text}
					invalid={editing.invalid}
					element={editorElement}
					onText={text => grid.setEditText(text)}
					onKeyDown={handleEditorKey}
					onBlur={handleEditorBlur}
				/>
			)}
			<div
				className="fo-sheet-live"
				role="status"
				aria-live="polite"
				aria-atomic="true"
			>
				{liveText}
			</div>
			<div
				ref={tipRef}
				className="fo-sheet-tip"
				role="status"
				hidden
			/>
		</div>
	)
}
