import { useSidebarWidthQuery } from "@/features/shell/queries/sidebarWidth"
import {
	defaultSidebarWidth,
	setSidebarWidth,
	widthFromDrag,
	widthFromKey,
	widthFromLeadingDrag,
	widthFromLeadingKey,
	SIDEBAR_EDGE,
	type SidebarEdge,
	type SidebarModule
} from "@/features/shell/lib/sidebarWidth"
import { useSeparatorValue, type SeparatorValue } from "@/lib/useSeparatorValue"

export type ResizableSidebarHandle = Omit<SeparatorValue, "value"> & { width: number; edge: SidebarEdge }

// Drive/Notes/Chats/Playlists/Assistant each call this with their own SidebarModule so the widths persist
// independently under features/shell/lib/sidebarWidth.ts's per-module kv keys.
export function useResizableSidebar(module: SidebarModule): ResizableSidebarHandle {
	const widthQuery = useSidebarWidthQuery(module)
	const edge = SIDEBAR_EDGE[module]
	const { value, ...handlers } = useSeparatorValue({
		persisted: widthQuery.data ?? defaultSidebarWidth(module),
		fromPointer: edge === "leading" ? widthFromLeadingDrag : widthFromDrag,
		fromKey: edge === "leading" ? widthFromLeadingKey : widthFromKey,
		commit: width => setSidebarWidth(module, width),
		refetch: () => widthQuery.refetch()
	})

	return { width: value, edge, ...handlers }
}
