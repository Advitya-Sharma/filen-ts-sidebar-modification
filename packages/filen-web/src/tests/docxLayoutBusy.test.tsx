// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from "vitest"
import { act, cleanup, render } from "@testing-library/react"
import { DocxLayoutBusy } from "@/features/preview/components/docxLayoutBusy"
import "@/lib/i18n"

afterEach(() => {
	cleanup()
})

// The browser starts the show animation once its 400 ms delay passed, and ends the fade-out: both by events.
function animation(element: Element, type: "animationstart" | "animationend"): void {
	act(() => {
		element.dispatchEvent(new Event(type, { bubbles: true }))
	})
}

describe("DocxLayoutBusy", () => {
	it("is a thin bar screen readers skip, shown after 400 ms by its animation alone, a static tint under reduced motion", () => {
		const { container } = render(<DocxLayoutBusy busy />)
		const bar = container.querySelector<HTMLElement>("[data-docx-layout-busy]")

		expect(bar).not.toBeNull()
		expect(bar?.getAttribute("aria-hidden")).toBe("true")
		expect(bar?.title).toBe("Laying out pages…")

		const classes = bar?.className.split(" ") ?? []

		for (const name of ["h-0.5", "animate-in", "fade-in", "delay-400", "fill-mode-backwards", "motion-reduce:bg-primary/40"]) {
			expect(classes).toContain(name)
		}

		const segment = bar?.firstElementChild?.className.split(" ") ?? []

		expect(segment).toContain("animate-indeterminate")
		expect(segment).toContain("motion-reduce:hidden")
	})

	it("never shows for a layout done within 400 ms, and stays a while once shown so it does not blink", () => {
		const view = render(<DocxLayoutBusy busy={false} />)
		const bar = (): HTMLElement | null => view.container.querySelector<HTMLElement>("[data-docx-layout-busy]")

		expect(bar()).toBeNull()

		// Done before its show animation started: gone at once.
		view.rerender(<DocxLayoutBusy busy />)
		expect(bar()).not.toBeNull()
		view.rerender(<DocxLayoutBusy busy={false} />)
		expect(bar()).toBeNull()

		// Shown, then done: it stays 300 ms and fades out, then goes.
		view.rerender(<DocxLayoutBusy busy />)

		const shown = bar()

		if (shown === null) throw new Error("no bar")

		animation(shown, "animationstart")
		view.rerender(<DocxLayoutBusy busy={false} />)
		expect(bar()?.dataset["docxLayoutBusy"]).toBe("leaving")
		expect(bar()?.className.split(" ")).toEqual(expect.arrayContaining(["animate-out", "fade-out", "delay-300", "fill-mode-forwards"]))

		// The segment's own animation is not the bar's.
		const segment = bar()?.firstElementChild

		if (segment === null || segment === undefined) throw new Error("no segment")

		animation(segment, "animationend")
		expect(bar()).not.toBeNull()

		// Busy again while fading: it stays on, without waiting 400 ms to show anew.
		view.rerender(<DocxLayoutBusy busy />)
		expect(bar()?.dataset["docxLayoutBusy"]).toBe("")
		expect(bar()?.className.split(" ")).not.toContain("animate-in")
		expect(bar()?.className.split(" ")).not.toContain("animate-out")

		view.rerender(<DocxLayoutBusy busy={false} />)

		const leaving = bar()

		if (leaving === null) throw new Error("no bar")

		animation(leaving, "animationend")
		expect(bar()).toBeNull()

		// The next layout waits its 400 ms again.
		view.rerender(<DocxLayoutBusy busy />)
		expect(bar()?.className.split(" ")).toContain("delay-400")
	})
})
