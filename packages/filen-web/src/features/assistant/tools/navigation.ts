import { useEffect } from "react"
import { useRouter, type RegisteredRouter } from "@tanstack/react-router"
import type { AssistantSignal } from "@filen/shared/assistant"
import { AssistantToolError } from "@/features/assistant/tools/format"
import { ASSISTANT_PANEL_ID, ASSISTANT_TOGGLE_ID, assistantInitialFocusTarget } from "@/features/assistant/components/assistantFocus"

// Tools run outside React, and the router instance is not exported (main.tsx), so a host mounted inside
// the router tree lends it to them: navigation goes through the same router.navigate the app's own links
// and menus use.
let mountedRouter: RegisteredRouter | null = null

export function setAssistantRouter(router: RegisteredRouter | null): void {
	mountedRouter = router
}

export function assistantRouter(): RegisteredRouter | null {
	return mountedRouter
}

// A navigation a dialog refused (an unsaved-changes prompt the user kept, a panel that blocks every
// navigation) is given up after this long.
const NAVIGATION_TIMEOUT_MS = 10_000

// How long after landing a destination may still take focus on its own: its mount effects, or a row it
// focuses once its listing loads.
const FOCUS_HOLD_MS = 5_000
// A destination that takes focus back every time it loses it is left alone after this many restores.
const FOCUS_RESTORE_LIMIT = 3
const DIALOG_SELECTOR = '[role="dialog"], [role="alertdialog"]'

let releaseHeldFocus: (() => void) | null = null

// Destinations focus themselves when they mount (a chat thread its composer, a revealed drive row), and the
// assistant's composer stays editable during a turn, so a user typing ahead would carry on in the chat
// composer, where their Enter sends. Focus the user left in the panel is put back when something takes it
// without the user's doing; clicking elsewhere, tabbing away or a dialog taking over ends the hold.
// Returns the release, which waits `afterMs` so effects of the landed destination are covered.
function holdAssistantFocus(): (afterMs: number) => void {
	releaseHeldFocus?.()

	if (typeof document === "undefined") {
		return () => undefined
	}

	const held = document.activeElement
	const outside = (node: EventTarget | null) => {
		const panel = document.getElementById(ASSISTANT_PANEL_ID)

		return panel === null || !(node instanceof Node) || !panel.contains(node)
	}

	if (!(held instanceof HTMLElement) || outside(held)) {
		return () => undefined
	}

	let restores = 0
	let userTask = false
	let userTaskTimer: ReturnType<typeof setTimeout> | undefined
	let releaseTimer: ReturnType<typeof setTimeout> | undefined
	const restore = () => {
		if (releaseHeldFocus !== release || !outside(document.activeElement)) {
			return
		}

		const target = held.isConnected ? held : assistantInitialFocusTarget()

		if (target === null || restores >= FOCUS_RESTORE_LIMIT) {
			release()

			return
		}

		restores += 1
		target.focus({ preventScroll: true })
	}
	// A focus change inside the task of a click or a key that acts (closing the panel, a menu, a dialog) is
	// the user's; typing keys are not, so a typed character never ends the hold.
	const markUserTask = () => {
		userTask = true
		clearTimeout(userTaskTimer)
		userTaskTimer = setTimeout(() => {
			userTask = false
		}, 0)
	}
	const onPointerDown = (event: PointerEvent) => {
		if (outside(event.target)) {
			release()

			return
		}

		markUserTask()
	}
	const onKeyDown = (event: KeyboardEvent) => {
		if (event.key === "Tab") {
			release()

			return
		}

		if (event.key === "Escape" || event.key === "Enter" || event.key === " " || event.ctrlKey || event.metaKey || event.altKey) {
			markUserTask()
		}
	}
	const onFocusIn = (event: FocusEvent) => {
		const target = event.target

		if (!outside(target)) {
			return
		}

		if (userTask || (target instanceof Element && (target.id === ASSISTANT_TOGGLE_ID || target.closest(DIALOG_SELECTOR) !== null))) {
			release()

			return
		}

		// Once the running script is done, so the destination's own focus handling (a caret placed after
		// focus()) has run.
		queueMicrotask(restore)
	}
	const release = () => {
		if (releaseHeldFocus !== release) {
			return
		}

		releaseHeldFocus = null
		clearTimeout(userTaskTimer)
		clearTimeout(releaseTimer)
		document.removeEventListener("pointerdown", onPointerDown, true)
		document.removeEventListener("keydown", onKeyDown, true)
		document.removeEventListener("focusin", onFocusIn, true)
	}

	releaseHeldFocus = release
	document.addEventListener("pointerdown", onPointerDown, true)
	document.addEventListener("keydown", onKeyDown, true)
	document.addEventListener("focusin", onFocusIn, true)

	return afterMs => {
		if (releaseHeldFocus !== release) {
			return
		}

		clearTimeout(releaseTimer)

		if (afterMs <= 0) {
			release()

			return
		}

		releaseTimer = setTimeout(release, afterMs)
	}
}

// Runs `navigate` and settles once the app shows where it went (`landed` reads the pathname), or fails when
// it did not get there. router.navigate's own promise settles only after a route load, and a blocked push
// starts none, so it would stay pending until some later navigation: the tool would hang its turn and report
// a screen that never opened. Stop ends the wait.
export function navigateChecked({
	router,
	navigate,
	landed,
	signal
}: {
	router: RegisteredRouter
	navigate: () => Promise<void>
	landed: (pathname: string) => boolean
	signal?: AssistantSignal | undefined
}): Promise<void> {
	if (signal?.aborted === true) {
		return Promise.reject(new AssistantToolError("Stopped"))
	}

	const releaseFocus = holdAssistantFocus()

	return new Promise((resolve, reject) => {
		let settled = false
		const arrived = () => landed(router.state.location.pathname)
		const blocked = () => new AssistantToolError("The app did not navigate: an open dialog kept the current screen")
		const finish = (outcome: () => void, arrivedThere: boolean) => {
			if (settled) {
				return
			}

			settled = true
			unsubscribe()
			clearTimeout(timer)
			signal?.removeEventListener("abort", onAbort)
			releaseFocus(arrivedThere ? FOCUS_HOLD_MS : 0)
			outcome()
		}
		const settle = () => {
			if (arrived()) {
				finish(resolve, true)

				return
			}

			finish(() => {
				reject(blocked())
			}, false)
		}
		const onAbort = () => {
			// A stopped navigation may still land, so its focus is held as long as a landed one's.
			finish(() => {
				reject(new AssistantToolError("Stopped"))
			}, true)
		}
		const unsubscribe = router.subscribe("onResolved", () => {
			if (arrived()) {
				finish(resolve, true)
			}
		})
		const timer = setTimeout(settle, NAVIGATION_TIMEOUT_MS)

		signal?.addEventListener("abort", onAbort, { once: true })
		navigate().then(settle, (e: unknown) => {
			finish(() => {
				reject(e instanceof Error ? e : new AssistantToolError(String(e)))
			}, false)
		})
	})
}

// For the assistant panel to call once, anywhere under the RouterProvider.
export function useAssistantNavigationBridge(): void {
	const router = useRouter()

	useEffect(() => {
		setAssistantRouter(router)

		return () => {
			if (mountedRouter === router) {
				setAssistantRouter(null)
			}
		}
	}, [router])
}
