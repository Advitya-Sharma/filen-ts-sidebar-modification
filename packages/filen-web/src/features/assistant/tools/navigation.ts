import { useEffect } from "react"
import { useRouter, type RegisteredRouter } from "@tanstack/react-router"

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
