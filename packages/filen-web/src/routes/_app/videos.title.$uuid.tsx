import { createFileRoute } from "@tanstack/react-router"
import { routeHead } from "@/lib/head/routeHead"
import { i18n } from "@/lib/i18n"
import { VideoTitleScreen } from "@/features/videos/screens/videoTitle"

// A title's page (`$uuid` = the title's id): a movie's or another video's details, or a series with its seasons.
export const Route = createFileRoute("/_app/videos/title/$uuid")({
	head: routeHead({ title: () => [i18n.t("videos:detailPageTitle"), i18n.t("common:moduleVideos")] }),
	component: TitlePage
})

function TitlePage() {
	const { uuid } = Route.useParams()

	// Keyed: another title is another page (its own scroll offset, dialog and follow), not this one's state.
	return (
		<VideoTitleScreen
			key={uuid}
			titleId={uuid}
		/>
	)
}
