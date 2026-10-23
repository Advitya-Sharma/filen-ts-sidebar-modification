import { useTranslation } from "react-i18next"
import { FileIcon } from "lucide-react"
import { Button } from "@/components/ui/button"
import {
	DropdownMenu,
	DropdownMenuContent,
	DropdownMenuItem,
	DropdownMenuSub,
	DropdownMenuSubContent,
	DropdownMenuSubTrigger,
	DropdownMenuTrigger
} from "@/components/ui/dropdown-menu"
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip"

export interface DocxFileMenuProps {
	// Offers Download as ▸ PDF; undefined when the file may not be downloaded.
	onDownloadPdf: (() => void) | undefined
}

// The Word document's file actions in its toolbar: Download as ▸ PDF (OpenDocument joins it later). Nothing to
// offer: not shown.
export function DocxFileMenu({ onDownloadPdf }: DocxFileMenuProps) {
	const { t } = useTranslation("preview")

	if (onDownloadPdf === undefined) {
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
									aria-label={t("previewDocxFileMenu")}
									className="shrink-0"
								>
									<FileIcon />
								</Button>
							}
						/>
					}
				/>
				<TooltipContent>{t("previewDocxFileMenu")}</TooltipContent>
			</Tooltip>
			<DropdownMenuContent
				align="start"
				data-preview-surface
			>
				<DropdownMenuSub>
					<DropdownMenuSubTrigger>{t("previewDocxDownloadAs")}</DropdownMenuSubTrigger>
					<DropdownMenuSubContent>
						<DropdownMenuItem onClick={onDownloadPdf}>{t("previewDocxDownloadPdf")}</DropdownMenuItem>
					</DropdownMenuSubContent>
				</DropdownMenuSub>
			</DropdownMenuContent>
		</DropdownMenu>
	)
}
