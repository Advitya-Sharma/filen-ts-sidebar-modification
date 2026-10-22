import { useTranslation } from "react-i18next"
import { SearchIcon } from "lucide-react"
import type { SheetGridState } from "@filen/office/react/sheet"
import { TooltipIconButton } from "@/components/ui/tooltipIconButton"
import { openSheetFind } from "@/features/spreadsheet/find/findState"

// The toolbar's Find button: opens the find bar (Mod+F). Finding changes nothing, so it stays enabled while the
// toolbar's edits are not (`disabled`).
export function FindButton({ grid }: { grid: SheetGridState; disabled: boolean }) {
	const { t } = useTranslation("preview")

	return (
		<TooltipIconButton
			label={t("previewSheetFindAction")}
			onClick={() => {
				openSheetFind(grid.model, "find")
			}}
		>
			<SearchIcon />
		</TooltipIconButton>
	)
}
