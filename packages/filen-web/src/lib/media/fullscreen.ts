import { useSyncExternalStore } from "react"

// Element fullscreen with the WebKit-prefixed fallback Safari still needs on some versions. A player
// fullscreens its own container, not the <video>, so its controls come along.

interface WebKitDocument {
	webkitFullscreenElement?: Element | null
	webkitFullscreenEnabled?: boolean
	webkitExitFullscreen?: () => void
}

interface WebKitElement {
	webkitRequestFullscreen?: () => void
}

const FULLSCREEN_EVENTS = ["fullscreenchange", "webkitfullscreenchange"] as const

function webkitDocument(): WebKitDocument {
	return document as Document & WebKitDocument
}

function fullscreenElement(): Element | null {
	return document.fullscreenElement ?? webkitDocument().webkitFullscreenElement ?? null
}

export function fullscreenSupported(): boolean {
	return typeof document !== "undefined" && (document.fullscreenEnabled || webkitDocument().webkitFullscreenEnabled === true)
}

export async function toggleFullscreen(element: HTMLElement): Promise<void> {
	if (fullscreenElement() === element) {
		if (typeof document.exitFullscreen === "function") {
			await document.exitFullscreen()
		} else {
			webkitDocument().webkitExitFullscreen?.()
		}

		return
	}

	if (typeof element.requestFullscreen === "function") {
		await element.requestFullscreen()
	} else {
		;(element as HTMLElement & WebKitElement).webkitRequestFullscreen?.()
	}
}

function subscribe(notify: () => void): () => void {
	for (const event of FULLSCREEN_EVENTS) {
		document.addEventListener(event, notify)
	}

	return () => {
		for (const event of FULLSCREEN_EVENTS) {
			document.removeEventListener(event, notify)
		}
	}
}

export function useIsFullscreen(element: Element | null): boolean {
	return useSyncExternalStore(
		subscribe,
		() => element !== null && fullscreenElement() === element,
		() => false
	)
}

interface WebKitVideoElement {
	webkitEnterFullscreen?: () => void
	webkitExitFullscreen?: () => void
	webkitDisplayingFullscreen?: boolean
}

interface LockableOrientation {
	lock?: ((orientation: "landscape") => Promise<void>) | undefined
}

// The video player's fullscreen: its container where element fullscreen exists, else (iPhone) the <video>
// alone in the system player, where the page's own controls and overlays are not shown. On a touch
// screen the picture is turned to landscape when the browser allows it.
export function videoFullscreenSupported(video: HTMLVideoElement | null): boolean {
	return (
		fullscreenSupported() ||
		(video !== null && typeof (video as HTMLVideoElement & WebKitVideoElement).webkitEnterFullscreen === "function")
	)
}

export function isVideoNativeFullscreen(video: HTMLVideoElement): boolean {
	return (video as HTMLVideoElement & WebKitVideoElement).webkitDisplayingFullscreen === true
}

export async function toggleVideoFullscreen(container: HTMLElement, video: HTMLVideoElement): Promise<void> {
	if (fullscreenSupported()) {
		const entering = fullscreenElement() !== container

		await toggleFullscreen(container)

		if (entering && window.matchMedia("(pointer: coarse)").matches) {
			try {
				const orientation: LockableOrientation = screen.orientation

				if (typeof orientation.lock === "function") {
					await orientation.lock("landscape")
				}
			} catch {
				// Not allowed here (desktop, or a browser without orientation lock).
			}
		}

		return
	}

	const webkitVideo = video as HTMLVideoElement & WebKitVideoElement

	if (webkitVideo.webkitDisplayingFullscreen === true) {
		webkitVideo.webkitExitFullscreen?.()
	} else {
		webkitVideo.webkitEnterFullscreen?.()
	}
}
