import type { AppContext } from "../fixture"

export const PLAYING: Partial<AppContext> = {
	screen: "audio",
	playing: {
		handle: "f17",
		playlist: "p2",
		paused: false,
		shuffle: false,
		repeat: false
	}
}

export function selected({ directory, items }: { directory: string; items: string[] }): Partial<AppContext> {
	return {
		screen: "drive",
		directory,
		selection: items
	}
}

export function inDirectory(directory: string): Partial<AppContext> {
	return {
		screen: "drive",
		directory
	}
}

// The notes screen with a note open in the editor.
export function openNote(note: string): Partial<AppContext> {
	return {
		screen: "notes",
		note
	}
}

// The chats screen with a chat open.
export function openChat(chat: string): Partial<AppContext> {
	return {
		screen: "chats",
		chat
	}
}
