import { useTranslation } from "react-i18next"
import { ChevronDownIcon, FolderOpenIcon, ImageIcon, LinkIcon, PlusIcon, StickyNoteIcon } from "lucide-react"
import type { SheetGridState } from "@filen/office/react/sheet"
import { Button } from "@/components/ui/button"
import {
	DropdownMenu,
	DropdownMenuContent,
	DropdownMenuItem,
	DropdownMenuSeparator,
	DropdownMenuTrigger
} from "@/components/ui/dropdown-menu"
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip"
import { InsertChartMenu } from "@/features/spreadsheet/charts/insertChartMenu"
import { requestImage, requestShortcut } from "@/features/spreadsheet/objects/insertRequests"
import { objectsEditable } from "@/features/spreadsheet/objects/objects.logic"

// The toolbar's Insert menu (workbooks only): a picture from this device or the drive, a chart, a link, a note. Each
// item goes where its own feature takes it: pictures to the objects feature, a chart to the charts feature, a link and
// a note through the same routes as Ctrl/Cmd+K and Shift+F2.
export function InsertMenu({ grid, disabled }: { grid: SheetGridState; disabled: boolean }) {
	const { t } = useTranslation("preview")
	const model = grid.model
	// The sheet's protection may lock its objects: pictures cannot go in then.
	const pictures = objectsEditable(grid.layout)

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
									aria-label={t("previewSheetInsertMenu")}
									className="shrink-0 gap-0.5 px-1.5"
								>
									<PlusIcon />
									<ChevronDownIcon className="size-3" />
								</Button>
							}
						/>
					}
				/>
				<TooltipContent>{t("previewSheetInsertMenu")}</TooltipContent>
			</Tooltip>
			<DropdownMenuContent data-preview-surface>
				<DropdownMenuItem
					disabled={!pictures}
					onClick={() => {
						requestImage(model, "file")
					}}
				>
					<ImageIcon />
					{t("previewSheetObjectImageFromFile")}
				</DropdownMenuItem>
				<DropdownMenuItem
					disabled={!pictures}
					onClick={() => {
						requestImage(model, "drive")
					}}
				>
					<FolderOpenIcon />
					{t("previewSheetObjectImageFromDrive")}
				</DropdownMenuItem>
				<InsertChartMenu
					model={model}
					disabled={disabled}
				/>
				<DropdownMenuSeparator />
				<DropdownMenuItem
					onClick={() => {
						requestShortcut(model, "hyperlink")
					}}
				>
					<LinkIcon />
					{t("previewSheetInsertLink")}
				</DropdownMenuItem>
				<DropdownMenuItem
					onClick={() => {
						requestShortcut(model, "note")
					}}
				>
					<StickyNoteIcon />
					{t("previewSheetInsertNote")}
				</DropdownMenuItem>
			</DropdownMenuContent>
		</DropdownMenu>
	)
}
