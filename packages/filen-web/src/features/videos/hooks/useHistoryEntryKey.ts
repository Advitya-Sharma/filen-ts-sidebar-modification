import { useRouterState } from "@tanstack/react-router"

// The current history entry's key: the same again when Back returns to it, a new one for every navigation.
export function useHistoryEntryKey(): string {
	return useRouterState({ select: state => state.location.state.__TSR_key ?? state.location.state.key ?? state.location.href })
}
