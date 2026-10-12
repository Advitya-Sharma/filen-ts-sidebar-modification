import { useSidebarWidthQuery } from "@/features/shell/queries/sidebarWidth"
import {
	defaultSidebarWidth,
	setSidebarWidth,
	widthFromDrag,
	widthFromKey,
	widthFromLeadingDrag,
	widthFromLeadingKey,
	SIDEBAR_EDGE,
	SIDEBAR_WIDTH_MAX,
	type SidebarEdge,
	type SidebarModule
} from "@/features/shell/lib/sidebarWidth"
import { useSeparatorValue, type SeparatorValue } from "@/lib/useSeparatorValue"

export type ResizableSidebarHandle = Omit<SeparatorValue, "value"> & { width: number; max: number; edge: SidebarEdge }

// Drive/Notes/Chats/Playlists/Assistant each call this with their own SidebarModule so the widths persist
// independently under features/shell/lib/sidebarWidth.ts's per-module kv keys. `max` is a live cap below
// SIDEBAR_WIDTH_MAX (the assistant's share of the viewport): the stored width stays as chosen, but the
// rendered width, the drag and key bases and the separator's value all use the capped one, so a width
// saved on a wider window has no dead zone on a narrower one.
export function useResizableSidebar(module: SidebarModule, max: number = SIDEBAR_WIDTH_MAX): ResizableSidebarHandle {
	const widthQuery = useSidebarWidthQuery(module)
	const edge = SIDEBAR_EDGE[module]
	const fromPointer = edge === "leading" ? widthFromLeadingDrag : widthFromDrag
	const fromKey = edge === "leading" ? widthFromLeadingKey : widthFromKey
	const { value, ...handlers } = useSeparatorValue({
		persisted: widthQuery.data ?? defaultSidebarWidth(module),
		fromPointer: (startWidth, startClientX, clientX) => Math.min(max, fromPointer(Math.min(max, startWidth), startClientX, clientX)),
		fromKey: (key, width) => {
			const next = fromKey(key, Math.min(max, width))

			return next === null ? null : Math.min(max, next)
		},
		commit: width => setSidebarWidth(module, width),
		refetch: () => widthQuery.refetch()
	})

	return { width: Math.min(max, value), max, edge, ...handlers }
}
