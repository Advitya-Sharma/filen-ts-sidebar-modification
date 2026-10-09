import { useTranslation } from "react-i18next"
import { MinusIcon, PlusIcon } from "lucide-react"
import { DENSITY_STEPS, DEFAULT_DENSITY_INDEX, clampDensityIndex } from "@/features/photos/lib/gridDensity"
import { setVideosGridDensity, useVideosGridDensityQuery } from "@/features/videos/queries/grid"
import { TooltipIconButton } from "@/components/ui/tooltipIconButton"

// The library grids' card-size steps, in the Videos header; every grid reads the same stored density.
export function VideosDensityControls() {
	const { t } = useTranslation("videos")
	const densityQuery = useVideosGridDensityQuery()
	const densityIndex = densityQuery.data ?? DEFAULT_DENSITY_INDEX

	async function handleDensityChange(nextIndex: number): Promise<void> {
		await setVideosGridDensity(clampDensityIndex(nextIndex))
		void densityQuery.refetch()
	}

	return (
		<div className="flex items-center gap-1 max-sm:hidden">
			<TooltipIconButton
				label={t("videosGridDensityDecrease")}
				disabled={densityIndex <= 0}
				onClick={() => {
					void handleDensityChange(densityIndex - 1)
				}}
			>
				<MinusIcon />
			</TooltipIconButton>
			<TooltipIconButton
				label={t("videosGridDensityIncrease")}
				disabled={densityIndex >= DENSITY_STEPS.length - 1}
				onClick={() => {
					void handleDensityChange(densityIndex + 1)
				}}
			>
				<PlusIcon />
			</TooltipIconButton>
		</div>
	)
}
