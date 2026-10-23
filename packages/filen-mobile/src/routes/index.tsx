import { Redirect } from "expo-router"
import { useIsAuthed } from "@/lib/auth"
import { useStartScreenHref } from "@/features/settings/hooks/useStartScreenHref"

export default function Index() {
	const isAuthed = useIsAuthed()
	const startScreenHref = useStartScreenHref()

	// Login and logout both reload the app, so this is the boot route either way.
	return <Redirect href={isAuthed ? startScreenHref : "/auth/login"} />
}
