import { createFileRoute } from "@tanstack/react-router"
import { routeHead } from "@/lib/head/routeHead"
import { i18n } from "@/lib/i18n"
import { VideosSearchScreen } from "@/features/videos/screens/videosSearch"

export const Route = createFileRoute("/_app/videos/search")({
	head: routeHead({ title: () => [i18n.t("videos:videosSearch"), i18n.t("common:moduleVideos")] }),
	component: VideosSearchScreen
})
