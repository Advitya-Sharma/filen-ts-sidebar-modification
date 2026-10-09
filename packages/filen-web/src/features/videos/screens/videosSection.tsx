// A library view (Movies, Shows, Other, Search, one directory) under the shared header: the place its grid is
// rendered into.
export function VideosSectionScreen() {
	return (
		<div
			data-slot="videos-section"
			className="flex min-h-0 flex-1 flex-col overflow-y-auto"
		/>
	)
}
