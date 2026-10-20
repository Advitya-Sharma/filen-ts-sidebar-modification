// @vitest-environment jsdom

import { useState } from "react"
import { describe, expect, it, vi } from "vitest"
import { fireEvent, render, screen, waitFor } from "@testing-library/react"
import { ContextMenu, ContextMenuContent, ContextMenuItem, ContextMenuTrigger } from "@/components/ui/context-menu"
import { Dialog, DialogContent, DialogTitle } from "@/components/ui/dialog"
import { useEscapeClosesMenu } from "@/features/spreadsheet/next/components/menuEscape"

// A menu over the grid inside the preview's dialog: Escape with focus still in the grid closes the menu, never the
// dialog under it.

function Viewer({ guarded, onDialog }: { guarded: boolean; onDialog: (open: boolean) => void }) {
	const [open, setOpen] = useState(false)

	useEscapeClosesMenu(guarded && open, () => {
		setOpen(false)
	})

	return (
		<Dialog
			open
			onOpenChange={next => {
				onDialog(next)
			}}
		>
			<DialogContent>
				<DialogTitle>Viewer</DialogTitle>
				<ContextMenu
					open={open}
					onOpenChange={setOpen}
				>
					<ContextMenuTrigger>
						<button type="button">grid</button>
					</ContextMenuTrigger>
					<ContextMenuContent>
						<ContextMenuItem>Insert row above</ContextMenuItem>
					</ContextMenuContent>
				</ContextMenu>
			</DialogContent>
		</Dialog>
	)
}

async function escapeWithGridFocused(guarded: boolean): Promise<(open: boolean) => void> {
	const onDialog = vi.fn()

	render(
		<Viewer
			guarded={guarded}
			onDialog={onDialog}
		/>
	)

	const grid = screen.getByRole("button", { name: "grid" })

	fireEvent.contextMenu(grid)
	await screen.findByRole("menuitem", { name: "Insert row above" })
	grid.focus()
	fireEvent.keyDown(grid, { key: "Escape", code: "Escape" })

	if (guarded) {
		await waitFor(() => {
			expect(screen.queryByRole("menuitem", { name: "Insert row above" })).toBeNull()
		})
	}

	return onDialog
}

describe("Escape over the grid's menu", () => {
	it("closes the menu and leaves the dialog open", async () => {
		const onDialog = await escapeWithGridFocused(true)

		expect(onDialog).not.toHaveBeenCalledWith(false)
	})

	it("closed the dialog too without the guard", async () => {
		const onDialog = await escapeWithGridFocused(false)

		expect(onDialog).toHaveBeenCalledWith(false)
	})
})
