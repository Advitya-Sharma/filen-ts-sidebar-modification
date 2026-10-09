// Track names shared by the player's tracks menu and the detail pages' selects.

export function channelLabel(channels: number | null, t: (key: "playerChannelsMono" | "playerChannelsStereo") => string): string | null {
	switch (channels) {
		case null:
			return null
		case 1:
			return t("playerChannelsMono")
		case 2:
			return t("playerChannelsStereo")
		case 6:
			return "5.1"
		case 8:
			return "7.1"
		default:
			return `${String(channels)} ch`
	}
}
