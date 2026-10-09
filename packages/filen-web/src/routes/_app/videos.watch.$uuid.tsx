import { createFileRoute } from "@tanstack/react-router"
import { routeHead } from "@/lib/head/routeHead"
import { i18n } from "@/lib/i18n"
import { WatchScreen, type WatchSearch } from "@/features/videos/components/player/watchScreen"

// The Videos player for a title (`$uuid` = the title's id): a full-screen stage over the app. `file` names the
// item to play (a stable uuid; an episode, a version or an extra), the title's own choice without it; `from=start`
// plays from the beginning instead of the resume point.
function validateSearch(search: Record<string, unknown>): WatchSearch {
	const file = search["file"]
	const from = search["from"]

	return {
		...(typeof file === "string" && file.length > 0 ? { file } : {}),
		...(from === "start" ? { from: "start" as const } : {})
	}
}

export const Route = createFileRoute("/_app/videos/watch/$uuid")({
	head: routeHead({ title: () => [i18n.t("videos:playerPageTitle"), i18n.t("common:moduleVideos")] }),
	validateSearch,
	component: WatchPage
})

function WatchPage() {
	const { uuid } = Route.useParams()
	const search = Route.useSearch()

	return (
		<WatchScreen
			titleId={uuid}
			search={search}
		/>
	)
}
