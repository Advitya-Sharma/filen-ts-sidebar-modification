// @vitest-environment happy-dom

import { afterEach, describe, expect, it } from "vitest"
import { paneAfter } from "@/features/spreadsheet/next/components/sheetPanes"

// F6 and Shift+F6 in the sheet view: a full cycle over its panes, as Excel's F6 switches between the worksheet and the
// panes around it.

afterEach(() => {
	document.body.replaceChildren()
})

function view(panel: boolean, selected: boolean): HTMLElement {
	const root = document.createElement("div")

	root.innerHTML =
		'<div data-sheet-pane="toolbar"><button>Undo</button><button>Bold</button></div>' +
		'<div data-sheet-pane="formula"><input aria-label="Name box"></div>' +
		'<div data-sheet-pane="grid"><div class="fo-sheet"><div class="fo-sheet-scroller" tabindex="0"></div>' +
		`<div class="fo-sheet-object" tabindex="${selected ? "0" : "-1"}"${selected ? " data-selected" : ""}></div></div></div>` +
		(panel ? '<div data-sheet-pane="panel"><aside><h2 tabindex="-1">Chart</h2><button>Close chart editor</button></aside></div>' : "") +
		'<div data-sheet-pane="tabs"><button>Add sheet</button></div>'
	document.body.append(root)

	// happy-dom lays nothing out: every element counts as shown.
	for (const element of root.querySelectorAll<HTMLElement>("*"))
		element.getClientRects = () => [new DOMRect(0, 0, 10, 10)] as unknown as DOMRectList

	return root
}

// An element by its accessible name, else its class, else its text.
function named(element: HTMLElement | null): string {
	if (element === null) {
		return "none"
	}

	return element.getAttribute("aria-label") ?? (element.className === "" ? element.textContent : element.className)
}

describe("the sheet view's panes", () => {
	it("cycles forward from the toolbar over the formula bar, the cells, the chart panel and the tabs, and round", () => {
		const root = view(true, false)
		const order: string[] = []
		let at: Element = root.querySelector("button") ?? root

		for (let k = 0; k < 6; k++) {
			const next = paneAfter(root, at, true)

			order.push(named(next))
			at = next ?? at
		}

		expect(order).toEqual(["Name box", "fo-sheet-scroller", "Close chart editor", "Add sheet", "Undo", "Name box"])
	})

	it("cycles back with Shift+F6, and lands on the selected object rather than the cells", () => {
		const root = view(false, true)
		const tabs = root.querySelector('[data-sheet-pane="tabs"] button')

		expect(named(paneAfter(root, tabs ?? root, false))).toBe("fo-sheet-object")

		const toolbar = root.querySelector('[data-sheet-pane="toolbar"] button')

		// Back from the toolbar: round to the tabs.
		expect(paneAfter(root, toolbar ?? root, false)?.textContent).toBe("Add sheet")
	})
})
