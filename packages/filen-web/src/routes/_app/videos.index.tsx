import { createFileRoute } from "@tanstack/react-router"
import { VideosHomeScreen } from "@/features/videos/screens/videosHome"

export const Route = createFileRoute("/_app/videos/")({ component: VideosHomeScreen })
