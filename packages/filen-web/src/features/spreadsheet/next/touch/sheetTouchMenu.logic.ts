// The menu keeps this far from the selection (px), and stays this far inside the grid. It goes below the selection
// past the touch handle's reach there.
const MENU_GAP = 10
const EDGE = 8
const HANDLE_REACH = 22

// Where a menu `width` by `height` goes for a selection shown at `box` in a grid `area` wide and high.
export function touchMenuPlace(
	box: { x: number; y: number; width: number; height: number },
	menu: { width: number; height: number },
	area: { width: number; height: number }
): { x: number; y: number } {
	const above = box.y - MENU_GAP - menu.height
	const below = box.y + box.height + MENU_GAP + HANDLE_REACH / 2
	const y = above >= EDGE ? above : Math.min(below, Math.max(EDGE, area.height - menu.height - EDGE))
	const x = Math.min(Math.max(EDGE, box.x + box.width / 2 - menu.width / 2), Math.max(EDGE, area.width - menu.width - EDGE))

	return { x, y }
}
