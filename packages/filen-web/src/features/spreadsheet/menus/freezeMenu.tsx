import { useTranslation } from "react-i18next"
import { ChevronDownIcon, SnowflakeIcon } from "lucide-react"
import type { SheetGridState } from "@filen/office/react/sheet"
import { Button } from "@/components/ui/button"
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from "@/components/ui/dropdown-menu"
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip"
import { FREEZE_ITEMS, freezeEnabled, freezeTo } from "@/features/spreadsheet/menus/structureCommands"

// The toolbar's freeze menu: panes at the active cell, the top row, the first column, or none.
export function FreezeMenu({ grid, disabled }: { grid: SheetGridState; disabled: boolean }) {
	const { t } = useTranslation("preview")

	return (
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
									aria-label={t("previewSheetFreeze")}
									className="shrink-0 gap-0.5 px-1.5"
								>
									<SnowflakeIcon />
									<ChevronDownIcon className="size-3" />
								</Button>
							}
						/>
					}
				/>
				<TooltipContent>{t("previewSheetFreeze")}</TooltipContent>
			</Tooltip>
			<DropdownMenuContent data-preview-surface>
				{FREEZE_ITEMS.map(item => (
					<DropdownMenuItem
						key={item.target}
						disabled={!freezeEnabled(grid, item.target)}
						onClick={() => {
							freezeTo(grid, item.target)
						}}
					>
						{t(item.key)}
					</DropdownMenuItem>
				))}
			</DropdownMenuContent>
		</DropdownMenu>
	)
}
