import { createFileRoute } from "@tanstack/react-router"
import { VideosSectionScreen } from "@/features/videos/screens/videosSection"

// One library directory's titles; `uuid` is the directory's.
export const Route = createFileRoute("/_app/videos/library/$uuid")({ component: VideosSectionScreen })
