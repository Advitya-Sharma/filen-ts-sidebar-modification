import { createContext, useContext, useEffect } from "react"

// The video stage's root element. Everything that pops up over the stage (menus, tooltips, panels) portals
// into it: in element fullscreen only the fullscreen element's own subtree is drawn.
export const VideoStageContainerContext = createContext<HTMLElement | null>(null)

export function useVideoStageContainer(): HTMLElement | null {
	return useContext(VideoStageContainerContext)
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
