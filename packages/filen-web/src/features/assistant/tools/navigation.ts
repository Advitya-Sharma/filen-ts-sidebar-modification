import { useEffect } from "react"
import { useRouter, type RegisteredRouter } from "@tanstack/react-router"
import type { AssistantSignal } from "@filen/shared/assistant"
import { AssistantToolError } from "@/features/assistant/tools/format"

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

	return new Promise((resolve, reject) => {
		let settled = false
		const arrived = () => landed(router.state.location.pathname)
		const blocked = () => new AssistantToolError("The app did not navigate: an open dialog kept the current screen")
		const finish = (outcome: () => void) => {
			if (settled) {
				return
			}

			settled = true
			unsubscribe()
			clearTimeout(timer)
			signal?.removeEventListener("abort", onAbort)
			outcome()
		}
		const onAbort = () => {
			finish(() => {
				reject(new AssistantToolError("Stopped"))
			})
		}
		const unsubscribe = router.subscribe("onResolved", () => {
			if (arrived()) {
				finish(resolve)
			}
		})
		const timer = setTimeout(() => {
			finish(
				arrived()
					? resolve
					: () => {
							reject(blocked())
						}
			)
		}, NAVIGATION_TIMEOUT_MS)

		signal?.addEventListener("abort", onAbort, { once: true })
		navigate().then(
			() => {
				finish(
					arrived()
						? resolve
						: () => {
								reject(blocked())
							}
				)
			},
			(e: unknown) => {
				finish(() => {
					reject(e instanceof Error ? e : new AssistantToolError(String(e)))
				})
			}
		)
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
