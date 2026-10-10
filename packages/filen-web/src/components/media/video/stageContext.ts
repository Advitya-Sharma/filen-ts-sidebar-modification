import { createContext, useContext, useEffect } from "react"

// The video stage's root element. Everything that pops up over the stage (menus, tooltips, panels) portals
// into it: in element fullscreen only the fullscreen element's own subtree is drawn.
export const VideoStageContainerContext = createContext<HTMLElement | null>(null)

export function useVideoStageContainer(): HTMLElement | null {
	return useContext(VideoStageContainerContext)
}

// How a popup of the stage (a menu, a popover) tells it that it shows: +1 while open, -1 when that ends.
export const VideoStagePopupContext = createContext<((delta: 1 | -1) => void) | null>(null)

// Keeps the stage's controls up while `open` holds. Follows the popup's real state, so a close by any path (a
// controlled prop, an unmount with the bar) ends it too.
export function useStagePopupOpen(open: boolean): void {
	const report = useContext(VideoStagePopupContext)

	useEffect(() => {
		if (!open || report === null) {
			return
		}

		report(1)

		return () => {
			report(-1)
		}
	}, [open, report])
}

// A card or panel of the stage: when it goes, focus goes back where it was before it opened (its trigger), or
// to the stage, so the player's keys keep working.
export function useRestoreStageFocus(): void {
	const container = useVideoStageContainer()

	useEffect(() => {
		const previous = document.activeElement

		return () => {
			requestAnimationFrame(() => {
				const active = document.activeElement

				if (active !== null && active !== document.body && container?.contains(active) === true) {
					return
				}

				const target =
					previous instanceof HTMLElement && previous.isConnected && container?.contains(previous) === true ? previous : container

				target?.focus({ preventScroll: true })
			})
		}
	}, [container])
}

// Popups over the picture: the controls' dark tone without any backdrop blur, which would recomposite the
// blurred region on every video frame.
export const STAGE_POPUP_TONE_CLASS = "dark bg-neutral-950/85 text-foreground shadow-lg ring-1 ring-white/15"
