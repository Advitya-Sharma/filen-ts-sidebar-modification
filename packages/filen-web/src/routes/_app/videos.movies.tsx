import { createFileRoute } from "@tanstack/react-router"
import { routeHead } from "@/lib/head/routeHead"
import { i18n } from "@/lib/i18n"
import { VideosSectionScreen } from "@/features/videos/screens/videosSection"

export const Route = createFileRoute("/_app/videos/movies")({
	head: routeHead({ title: () => [i18n.t("videos:videosMovies"), i18n.t("common:moduleVideos")] }),
	component: VideosSectionScreen
})
