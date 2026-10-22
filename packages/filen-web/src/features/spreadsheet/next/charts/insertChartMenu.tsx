import { useTranslation } from "react-i18next"
import { ChartColumnIcon } from "lucide-react"
import type { SheetGridModel } from "@filen/office/react/sheet"
import { DropdownMenuItem, DropdownMenuSub, DropdownMenuSubContent, DropdownMenuSubTrigger } from "@/components/ui/dropdown-menu"
import { CHART_TYPES } from "@/features/spreadsheet/next/charts/chartPanel.logic"
import { requestChart } from "@/features/spreadsheet/next/charts/chartRequests"

// The Insert menu's Chart submenu: a chart of the selected cells, of the type picked (the charts feature inserts it).
export function InsertChartMenu(props: { model: SheetGridModel; disabled: boolean }) {
	const { t } = useTranslation("preview")

	return (
		<DropdownMenuSub>
			<DropdownMenuSubTrigger disabled={props.disabled}>
				<ChartColumnIcon />
				{t("previewSheetChartInsert")}
			</DropdownMenuSubTrigger>
			<DropdownMenuSubContent>
				{CHART_TYPES.map(type => (
					<DropdownMenuItem
						key={type.value}
						onClick={() => {
							requestChart(props.model, type.value)
						}}
					>
						{t(type.key)}
					</DropdownMenuItem>
				))}
			</DropdownMenuSubContent>
		</DropdownMenuSub>
	)
}
