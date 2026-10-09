import type { ComponentType } from "react"
import { useTranslation } from "react-i18next"
import { Link } from "@tanstack/react-router"
import { FilmIcon, HouseIcon, SearchIcon, TvIcon } from "lucide-react"

type IconType = ComponentType<{ className?: string }>

const TAB_CLASS =
	"flex h-14 flex-col items-center justify-center gap-0.5 rounded-xl text-[11px] font-medium text-muted-foreground focus-ring outline-none transition-colors data-[status=active]:text-foreground [&_svg]:size-5 [&_svg]:shrink-0"

function Tab({
	to,
	icon: Icon,
	label,
	exact
}: {
	to: "/videos" | "/videos/movies" | "/videos/shows" | "/videos/search"
	icon: IconType
	label: string
	exact?: boolean
}) {
	// TanStack stamps aria-current="page" and data-status="active" on the matching Link.
	return (
		<li>
			<Link
				to={to}
				activeOptions={{ exact: exact === true }}
				className={TAB_CLASS}
			>
				<Icon />
				<span>{label}</span>
			</Link>
		</li>
	)
}

// The phone tabs: the library's daily navigation at the bottom of the content panel, below the layout
// breakpoint only (pure CSS). The directories stay in the sidebar drawer; the audio bar keeps its own place
// below the panel, and the panel's own inset keeps the tabs clear of the screen edge.
export function VideosTabs() {
	const { t } = useTranslation(["videos", "common"])

	return (
		<nav
			aria-label={t("common:moduleVideos")}
			className="shrink-0 border-t border-border bg-card md:hidden"
		>
			<ul className="grid grid-cols-4 px-1">
				<Tab
					to="/videos"
					icon={HouseIcon}
					label={t("videosHome")}
					exact
				/>
				<Tab
					to="/videos/movies"
					icon={FilmIcon}
					label={t("videosMovies")}
				/>
				<Tab
					to="/videos/shows"
					icon={TvIcon}
					label={t("videosShows")}
				/>
				<Tab
					to="/videos/search"
					icon={SearchIcon}
					label={t("videosSearch")}
				/>
			</ul>
		</nav>
	)
}
