import { useTranslation } from "react-i18next"
import { ChevronDownIcon, TableCellsMergeIcon } from "lucide-react"
import type { SheetGridState } from "@filen/office/react/sheet"
import { Button } from "@/components/ui/button"
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from "@/components/ui/dropdown-menu"
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip"
import { MERGE_ITEMS, selectionHasMerge } from "@/features/spreadsheet/menus/structureCommands"
import { useMergeCommands } from "@/features/spreadsheet/menus/useMergeCommands"

// The toolbar's merge menu: merge and center, merge across, merge, unmerge.
export function MergeMenu({ grid, disabled }: { grid: SheetGridState; disabled: boolean }) {
	const { t } = useTranslation("preview")
	const { merge, unmerge, dialog } = useMergeCommands(grid)

	return (
		<>
			<DropdownMenu>
				<Tooltip>
					<TooltipTrigger
						render={
							<DropdownMenuTrigger
								render={
									<Button
										variant="ghost"
										size="sm"
										disabled={disabled}
										aria-label={t("previewSheetMerge")}
										className="shrink-0 gap-0.5 px-1.5"
									>
										<TableCellsMergeIcon />
										<ChevronDownIcon className="size-3" />
									</Button>
								}
							/>
						}
					/>
					<TooltipContent>{t("previewSheetMerge")}</TooltipContent>
				</Tooltip>
				<DropdownMenuContent data-preview-surface>
					{MERGE_ITEMS.map(item => (
						<DropdownMenuItem
							key={item.mode}
							onClick={() => {
								merge(item.mode)
							}}
						>
							{t(item.key)}
						</DropdownMenuItem>
					))}
					<DropdownMenuItem
						disabled={!selectionHasMerge(grid)}
						onClick={unmerge}
					>
						{t("previewSheetUnmerge")}
					</DropdownMenuItem>
				</DropdownMenuContent>
			</DropdownMenu>
			{dialog}
		</>
	)
}
