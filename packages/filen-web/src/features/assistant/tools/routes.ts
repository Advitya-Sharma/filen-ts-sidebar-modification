import type { RegisteredRouter } from "@tanstack/react-router"
import { openRoute } from "@filen/shared/assistant"
import { DEFAULT_CONTACTS_SECTION_FILTER } from "@/features/contacts/components/contactsList.logic"
import { pathIsUnder } from "@/features/shell/lib/appShell.logic"

export type AssistantRouteName = (typeof openRoute.input.shape.route.options)[number]

// Where each open_route section lives, most specific first so a settings page beats "settings". The
// web app has no camera upload, so that section has no path and opens nothing.
const ROUTE_PATHS: readonly (readonly [AssistantRouteName, string])[] = [
	["drive", "/drive"],
	["recents", "/recents"],
	["favorites", "/favorites"],
	["shared_with_me", "/shared-in"],
	["shared_by_me", "/shared-out"],
	["public_links", "/links"],
	["trash", "/trash"],
	["photos", "/photos"],
	["notes", "/notes"],
	["chats", "/chats"],
	["contacts", "/contacts"],
	["audio", "/playlists"],
	["transfers", "/transfers"],
	["settings_account", "/settings/account"],
	["settings_security", "/settings/security"],
	["settings_appearance", "/settings/appearance"],
	["settings", "/settings"]
]

// The section a pathname belongs to, in open_route's own words; "other" for anything else.
export function screenOfPath(pathname: string): AssistantRouteName | "other" {
	return ROUTE_PATHS.find(([, root]) => pathIsUnder(pathname, root))?.[0] ?? "other"
}

// The same navigations the icon rail and the settings sidebar make. False when the section does not
// exist on the web.
export async function navigateToRoute(router: RegisteredRouter, route: AssistantRouteName): Promise<boolean> {
	switch (route) {
		case "drive":
			await router.navigate({ to: "/drive/$", params: { _splat: "" } })
			break
		case "recents":
			await router.navigate({ to: "/recents" })
			break
		case "favorites":
			await router.navigate({ to: "/favorites" })
			break
		case "shared_with_me":
			await router.navigate({ to: "/shared-in/$", params: { _splat: "" } })
			break
		case "shared_by_me":
			await router.navigate({ to: "/shared-out/$", params: { _splat: "" } })
			break
		case "public_links":
			await router.navigate({ to: "/links" })
			break
		case "trash":
			await router.navigate({ to: "/trash" })
			break
		case "photos":
			await router.navigate({ to: "/photos" })
			break
		case "notes":
			await router.navigate({ to: "/notes" })
			break
		case "chats":
			await router.navigate({ to: "/chats" })
			break
		case "contacts":
			await router.navigate({ to: "/contacts", search: { section: DEFAULT_CONTACTS_SECTION_FILTER } })
			break
		case "audio":
			await router.navigate({ to: "/playlists", search: {} })
			break
		case "transfers":
			await router.navigate({ to: "/transfers" })
			break
		case "settings":
			await router.navigate({ to: "/settings" })
			break
		case "settings_account":
			await router.navigate({ to: "/settings/account" })
			break
		case "settings_security":
			await router.navigate({ to: "/settings/security" })
			break
		case "settings_appearance":
			await router.navigate({ to: "/settings/appearance" })
			break
		case "settings_camera_upload":
			return false
	}

	return true
}
