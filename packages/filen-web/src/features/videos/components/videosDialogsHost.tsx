import { useVideosDialogStore } from "@/features/videos/store/useVideosDialogStore"
import { AddDirectoryDialog } from "@/features/videos/components/addDirectoryDialog"
import { LibrariesDialog } from "@/features/videos/components/librariesDialog"

// The Videos dialogs, mounted only while open. The chooser can open over the Libraries dialog.
export function VideosDialogsHost() {
	const chooser = useVideosDialogStore(state => state.chooser)
	const libraries = useVideosDialogStore(state => state.libraries)

	return (
		<>
			{libraries ? <LibrariesDialog /> : null}
			{chooser === null ? null : <AddDirectoryDialog replacing={chooser.replacing} />}
		</>
	)
}
