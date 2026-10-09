import { create } from "zustand"

interface VideosDialogState {
	// The directory chooser, with the library directories the pick replaces (an empty one being swapped).
	chooser: { replacing: string[] } | null
	libraries: boolean
}

// The Videos dialogs. A store rather than component state because the sidebar (the shell's) and the screens
// (the route's) both open them; VideosDialogsHost, mounted by the Videos layout, renders them while open.
export const useVideosDialogStore = create<VideosDialogState>(() => ({
	chooser: null,
	libraries: false
}))

export function openVideosChooser(replacing: string[] = []): void {
	useVideosDialogStore.setState({
		chooser: {
			replacing
		}
	})
}

export function closeVideosChooser(): void {
	useVideosDialogStore.setState({
		chooser: null
	})
}

export function setVideosLibrariesOpen(open: boolean): void {
	useVideosDialogStore.setState({
		libraries: open
	})
}

export function closeVideosDialogs(): void {
	useVideosDialogStore.setState({
		chooser: null,
		libraries: false
	})
}
