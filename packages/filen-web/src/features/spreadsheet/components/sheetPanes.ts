// The sheet view's panes, in the order F6 visits them (Excel's F6 "switches between the worksheet, ribbon, task pane"
// and Shift+F6 goes the other way): the toolbar, the formula bar, the cells (or the object selected on them), the chart
// editor when it is open, and the sheet tabs. Each is an element marked data-sheet-pane; the cycle wraps round.

const TABBABLE = 'a[href], button, input, select, textarea, [tabindex]:not([tabindex="-1"]), [contenteditable="true"]'

// Where focus goes in a pane: the selected object or the cells for the grid, else the first control that takes Tab.
function landing(pane: Element): HTMLElement | null {
	if (pane.getAttribute("data-sheet-pane") === "grid") {
		return pane.querySelector<HTMLElement>(".fo-sheet-object[data-selected]") ?? pane.querySelector<HTMLElement>(".fo-sheet-scroller")
	}

	for (const element of pane.querySelectorAll<HTMLElement>(TABBABLE)) {
		if (!element.matches(":disabled") && element.getClientRects().length > 0) {
			return element
		}
	}

	return null
}

// The element focus goes to from `from` with F6 (forward) or Shift+F6, among the panes in `root`; null when none takes it.
export function paneAfter(root: Element, from: Element, forward: boolean): HTMLElement | null {
	const panes = Array.from(root.querySelectorAll("[data-sheet-pane]"))
	const current = panes.findIndex(pane => pane.contains(from))
	const step = forward ? 1 : -1
	// Not in a pane: forward starts at the first, back at the last.
	const start = current < 0 ? (forward ? -1 : panes.length) : current

	for (let k = 1; k <= panes.length; k++) {
		const pane = panes[(((start + step * k) % panes.length) + panes.length) % panes.length]
		const target = pane === undefined || pane === panes[current] ? null : landing(pane)

		if (target !== null) {
			return target
		}
	}

	return null
}
