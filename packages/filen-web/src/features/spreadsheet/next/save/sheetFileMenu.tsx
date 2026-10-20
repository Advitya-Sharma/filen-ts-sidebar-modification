import { useTranslation } from "react-i18next"
import { toast } from "sonner"
import { FileSpreadsheetIcon } from "lucide-react"
import { Button } from "@/components/ui/button"
import {
	DropdownMenu,
	DropdownMenuCheckboxItem,
	DropdownMenuContent,
	DropdownMenuItem,
	DropdownMenuSeparator,
	DropdownMenuTrigger
} from "@/components/ui/dropdown-menu"
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip"
import { saveOfficeAutosave, useOfficeSettings } from "@/features/office/lib/officeSettings"
import { useIsOnline } from "@/lib/useIsOnline"

export interface SheetFileMenuProps {
	// A text file (its CSV is its cells) rather than a workbook (its CSV is the sheet on show).
	text: boolean
	// Offers Download as CSV (never for a password-protected workbook, whose cells stay sealed).
	onDownloadCsv: (() => void) | undefined
	// Offers Save as .xlsx (asks first).
	onSaveAsXlsx: (() => void) | undefined
	// Offers the autosave setting: the file is edited here.
	autosave: boolean
}

// The spreadsheet's file actions beside the cell contents: Save as .xlsx, Download as CSV, and the autosave setting
// (the same synced setting as Settings → Advanced → Documents). Nothing to offer: not shown.
export function SheetFileMenu({ text, onDownloadCsv, onSaveAsXlsx, autosave }: SheetFileMenuProps) {
	const { t } = useTranslation(["preview", "settings"])
	const settings = useOfficeSettings()
	const online = useIsOnline()

	if (onDownloadCsv === undefined && onSaveAsXlsx === undefined && !autosave) {
		return null
	}

	return (
		<DropdownMenu>
			<Tooltip>
				<TooltipTrigger
					render={
						<DropdownMenuTrigger
							render={
								<Button
									variant="ghost"
									size="icon-xs"
									aria-label={t("previewSheetFileMenu")}
									className="shrink-0"
								>
									<FileSpreadsheetIcon />
								</Button>
							}
						/>
					}
				/>
				<TooltipContent>{t("previewSheetFileMenu")}</TooltipContent>
			</Tooltip>
			<DropdownMenuContent
				align="end"
				data-preview-surface
			>
				{onSaveAsXlsx === undefined ? null : (
					<DropdownMenuItem onClick={onSaveAsXlsx}>{t("previewSpreadsheetSaveAsXlsx")}</DropdownMenuItem>
				)}
				{onDownloadCsv === undefined ? null : (
					<DropdownMenuItem onClick={onDownloadCsv}>
						{t(text ? "previewSheetDownloadCsv" : "previewSheetDownloadSheetCsv")}
					</DropdownMenuItem>
				)}
				{autosave ? (
					<>
						{onSaveAsXlsx === undefined && onDownloadCsv === undefined ? null : <DropdownMenuSeparator />}
						<DropdownMenuCheckboxItem
							closeOnClick
							checked={settings.autosave}
							disabled={!settings.loaded || !online}
							onCheckedChange={checked => {
								saveOfficeAutosave(checked).catch(() => {
									toast.error(t("settings:settingsDocumentsSaveFailed"))
								})
							}}
						>
							{t("previewSheetAutosave")}
						</DropdownMenuCheckboxItem>
					</>
				) : null}
			</DropdownMenuContent>
		</DropdownMenu>
	)
}
