import { useLayoutEffect, useRef, useState, type KeyboardEvent, type RefObject } from "react"
import { useTranslation } from "react-i18next"
import type { SheetRange } from "@filen/office/core/sheet"
import type { SheetGridHandle } from "@filen/office/react/sheet"
import { cn } from "@filen/shared"
import { touchMenuPlace } from "@/features/spreadsheet/touch/sheetTouchMenu.logic"

// The touch menu over the selection, as the DOCX editor's: a compact row of Cut, Copy, Paste and Select all (and
// Edit, which opens the active cell with the soft keyboard), above the selection or below it when there is no room.
// It sits in the grid's overlay layer and re-places itself on every frame the grid draws, without a render. A row of
// buttons, so a toolbar (WAI-ARIA APG): one Tab stop, the arrows, Home and End move between them.

export interface SheetTouchActions {
	edit: (() => void) | null
	cut: (() => void) | null
	copy: () => void
	paste: (() => void) | null
	selectAll: () => void
	// Opens the grid's full menu (insert, delete, hide, merge, notes, links, sort, fill, freeze, objects, charts).
	more: () => void
}

export function SheetTouchMenu({
	gridRef,
	range,
	actions
}: {
	gridRef: RefObject<SheetGridHandle | null>
	// The selection it is about (its last range).
	range: SheetRange
	actions: SheetTouchActions
}) {
	const { t } = useTranslation(["preview", "common"])
	const menuRef = useRef<HTMLDivElement>(null)
	// The button the toolbar's Tab stop is on.
	const [current, setCurrent] = useState(0)

	useLayoutEffect(() => {
		const handle = gridRef.current

		if (handle === null) {
			return undefined
		}

		const place = (): void => {
			const menu = menuRef.current
			const area = menu?.parentElement

			if (menu === null || area === null || area === undefined) {
				return
			}

			const box = handle.rangeRect(range)

			// Scrolled away: hidden until it is back.
			if (box === null) {
				menu.style.visibility = "hidden"

				return
			}

			const at = touchMenuPlace(
				box,
				{ width: menu.offsetWidth, height: menu.offsetHeight },
				{ width: area.clientWidth, height: area.clientHeight }
			)

			menu.style.visibility = ""
			menu.style.transform = `translate(${String(at.x)}px, ${String(at.y)}px)`
		}

		place()

		return handle.onFrame(place)
	}, [gridRef, range])

	const items = [
		{ key: "cut", label: t("previewSheetTouchCut"), run: actions.cut },
		{ key: "copy", label: t("previewSheetTouchCopy"), run: actions.copy },
		{ key: "paste", label: t("previewSheetTouchPaste"), run: actions.paste },
		{ key: "selectAll", label: t("previewSheetTouchSelectAll"), run: actions.selectAll },
		{ key: "edit", label: t("previewSheetTouchEdit"), run: actions.edit },
		{ key: "more", label: "⋯", name: t("previewSheetTouchMore"), run: actions.more }
	].filter(item => item.run !== null)
	const stop = Math.min(current, items.length - 1)

	function handleKeyDown(event: KeyboardEvent<HTMLDivElement>): void {
		const last = items.length - 1
		const next =
			event.key === "ArrowRight"
				? (stop + 1) % items.length
				: event.key === "ArrowLeft"
					? (stop + last) % items.length
					: event.key === "Home"
						? 0
						: event.key === "End"
							? last
							: null

		if (next === null) {
			return
		}

		event.preventDefault()
		setCurrent(next)
		menuRef.current?.querySelectorAll<HTMLButtonElement>("button")[next]?.focus()
	}

	return (
		<div
			ref={menuRef}
			role="toolbar"
			aria-label={t("common:selectionActionsLabel")}
			data-preview-surface
			className="pointer-events-auto absolute top-0 left-0 flex h-11 items-center overflow-hidden rounded-xl bg-popover text-sm text-popover-foreground shadow-md ring-1 ring-foreground/10"
			style={{ visibility: "hidden" }}
			// Pressing the menu keeps the grid's focus and selection.
			onPointerDown={event => {
				event.preventDefault()
			}}
			onKeyDown={handleKeyDown}
		>
			{items.map((item, index) => (
				<button
					key={item.key}
					type="button"
					tabIndex={index === stop ? 0 : -1}
					aria-label={"name" in item ? item.name : undefined}
					className={cn("h-full min-w-11 px-3 whitespace-nowrap active:bg-accent", index > 0 && "border-l border-border")}
					onFocus={() => {
						setCurrent(index)
					}}
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
