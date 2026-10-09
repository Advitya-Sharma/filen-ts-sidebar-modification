// @vitest-environment jsdom

import { afterEach, describe, expect, it, vi } from "vitest"
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react"
import "@/lib/i18n"
import { SidebarDrawer, SidebarDrawerTrigger } from "@/features/shell/components/sidebarDrawer"

afterEach(() => {
	cleanup()
})

function shell(open: boolean, narrow: boolean, onOpenChange: (open: boolean) => void = () => undefined) {
	return (
		<>
			<SidebarDrawerTrigger />
			<SidebarDrawer
				open={open}
				narrow={narrow}
				label="Drive"
				panel={<p>Drive panel</p>}
				onOpenChange={onOpenChange}
			/>
		</>
	)
}

function settle(): Promise<void> {
	return act(() => new Promise<void>(resolve => setTimeout(resolve, 50)))
}

describe("SidebarDrawer", () => {
	it("leaves nothing behind when the window widens with the drawer open and narrows again", async () => {
		const { rerender } = render(shell(true, true))

		expect(await screen.findByRole("dialog", { name: "Drive" })).toBeTruthy()

		// The shell closes the drawer on the crossing, in the same render that drops its narrow layout.
		rerender(shell(false, false))
		rerender(shell(false, true))
		await settle()

		expect(screen.queryByRole("dialog", { name: "Drive" })).toBeNull()
		expect(document.querySelector("[data-ending-style]")).toBeNull()
		// Still mounted behind the closed drawer, so the panel keeps its state.
		expect(screen.getByText("Drive panel", { selector: "p" })).toBeTruthy()
	})

	it("opens from the rail's trigger again after a trip across the breakpoint", async () => {
		const onOpenChange = vi.fn()
		const { rerender } = render(shell(false, true, onOpenChange))

		rerender(shell(false, false, onOpenChange))
		rerender(shell(false, true, onOpenChange))
		await settle()
		fireEvent.click(screen.getByRole("button", { name: "Open navigation" }))

		expect(onOpenChange).toHaveBeenLastCalledWith(true, expect.anything())
	})
})
