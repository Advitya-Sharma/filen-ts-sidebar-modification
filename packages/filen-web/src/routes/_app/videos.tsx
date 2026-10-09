import { createFileRoute } from "@tanstack/react-router"
import { routeHead } from "@/lib/head/routeHead"
import { i18n } from "@/lib/i18n"
import { VideosLayout } from "@/features/videos/components/videosLayout"

// Videos layout parent: every /videos route renders inside it (the sidebar lives in the shell).
export const Route = createFileRoute("/_app/videos")({
	head: routeHead({ title: () => [i18n.t("common:moduleVideos")] }),
	component: VideosLayout
})
