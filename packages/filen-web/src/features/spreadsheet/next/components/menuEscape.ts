import { useEffect } from "react"

// While a menu over the grid is open, Escape closes the menu and nothing else. Base UI's menu and the preview's
// dialog both hear Escape at the document; with focus still in the grid (a menu that just opened, or one opened
// while focus went back to the grid) both act, and the viewer closed with the menu. Taken first, in the window's capture
// phase, Escape reaches neither: the menu is closed here.
export function useEscapeClosesMenu(open: boolean, close: () => void): void {
	useEffect(() => {
		if (!open) {
			return undefined
		}

		const onKeyDown = (event: KeyboardEvent): void => {
			if (event.key !== "Escape" || event.isComposing) {
				return
			}

			event.preventDefault()
			event.stopImmediatePropagation()
			close()
		}

		window.addEventListener("keydown", onKeyDown, true)

		return () => {
			window.removeEventListener("keydown", onKeyDown, true)
		}
	}, [open, close])
}
