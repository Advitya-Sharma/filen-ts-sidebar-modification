import type { ReactNode } from "react"
import { useTranslation } from "react-i18next"
import { FileDiffIcon } from "lucide-react"
import { isWriteCommand, type DocxEditorCommand, type DocxEditorCommandType } from "@filen/office/core/docx/edit"
import { DOCX_MARKUP_MODES, type DocxMarkupMode } from "@filen/office/dom/docx"
import type { DocxToolbarSlotProps } from "@/features/preview/components/docxEditor.logic"
import { MARKUP_LABELS } from "@/features/preview/components/docxViewer.logic"
import { Button } from "@/components/ui/button"
import {
	DropdownMenu,
	DropdownMenuCheckboxItem,
	DropdownMenuContent,
	DropdownMenuGroup,
	DropdownMenuItem,
	DropdownMenuLabel,
	DropdownMenuRadioGroup,
	DropdownMenuRadioItem,
	DropdownMenuSeparator,
	DropdownMenuTrigger
} from "@/components/ui/dropdown-menu"
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip"

export interface DocxReviewMenuProps extends DocxToolbarSlotProps {
	// Null for a document without comments or tracked changes.
	markup: DocxMarkupMode | null
	onMarkupChange: (mode: DocxMarkupMode) => void
}

// The menu's sections, in Word's Review tab order: what each runs (a command, or null: the input it asks for).
interface Item {
	readonly type: DocxEditorCommandType
	readonly command: DocxEditorCommand | null
}

const COMMENT_ITEMS: readonly Item[] = [
	{ type: "comment.add", command: null },
	{ type: "comment.prev", command: { type: "comment.prev" } },
	{ type: "comment.next", command: { type: "comment.next" } }
]
const CHANGE_ITEMS: readonly Item[] = [
	{ type: "review.prev", command: { type: "review.prev" } },
	{ type: "review.next", command: { type: "review.next" } },
	{ type: "review.accept", command: { type: "review.accept" } },
	{ type: "review.reject", command: { type: "review.reject" } }
]
const ALL_ITEMS: readonly Item[] = [
	{ type: "review.acceptAll", command: { type: "review.acceptAll" } },
	{ type: "review.rejectAll", command: { type: "review.rejectAll" } }
]

// The toolbar's one Review menu: the track changes switch (a dot on the button while on), a new comment and
// stepping through comments, stepping through changes and accepting or rejecting the one at the caret or all of
// them, and the markup modes (it replaces the markup menu).
export function DocxReviewMenu({ commands, snapshot, readOnly, execute, requestInput, markup, onMarkupChange }: DocxReviewMenuProps) {
	const { t } = useTranslation("preview")
	const tracking = snapshot?.review.tracking === true
	const labels = new Map(commands.map(command => [command.info.type, command.label]))
	const trackLabel = labels.get("review.track")
	const disabled = (type: DocxEditorCommandType): boolean =>
		snapshot === undefined ||
		!snapshot.enabled[type] ||
		(readOnly && isWriteCommand(type)) ||
		(snapshot.readOnly && isWriteCommand(type))
	const items = (list: readonly Item[]): ReactNode[] =>
		list.flatMap(({ type, command }) => {
			const label = labels.get(type)

			if (label === undefined) {
				return []
			}

			// Accepting or rejecting the change at the caret needs one there.
			const off = disabled(type) || ((type === "review.accept" || type === "review.reject") && snapshot?.review.current == null)

			return [
				<DropdownMenuItem
					key={type}
					disabled={off}
					onClick={() => {
						if (command === null) {
							requestInput(type)
						} else {
							execute(command)
						}
					}}
				>
					{t(label)}
				</DropdownMenuItem>
			]
		})
	const comments = items(COMMENT_ITEMS)
	const changes = items(CHANGE_ITEMS)
	const all = items(ALL_ITEMS)

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
									aria-label={tracking ? t("previewDocxReviewMenuTracking") : t("previewDocxReviewMenu")}
									className="relative px-2"
								>
									<FileDiffIcon />
									<span className="text-xs max-md:hidden">{t("previewDocxReviewMenu")}</span>
									{tracking ? (
										<span
											aria-hidden="true"
											data-tracking=""
											className="absolute top-1 right-1 size-1.5 rounded-full bg-destructive"
										/>
									) : null}
								</Button>
							}
						/>
					}
				/>
				<TooltipContent>{tracking ? t("previewDocxReviewMenuTracking") : t("previewDocxReviewMenu")}</TooltipContent>
			</Tooltip>
			<DropdownMenuContent
				align="end"
				className="w-auto min-w-48"
			>
				{trackLabel === undefined ? null : (
					<DropdownMenuCheckboxItem
						checked={tracking}
						closeOnClick
						disabled={disabled("review.track")}
						onCheckedChange={on => {
							execute({ type: "review.track", on })
						}}
					>
						{t(trackLabel)}
					</DropdownMenuCheckboxItem>
				)}
				{[comments, changes, all].map((group, k) =>
					group.length === 0 ? null : (
						<DropdownMenuGroup key={k}>
							<DropdownMenuSeparator />
							{group}
						</DropdownMenuGroup>
					)
				)}
				{markup === null ? null : (
					<>
						<DropdownMenuSeparator />
						<DropdownMenuRadioGroup
							value={markup}
							onValueChange={(mode: DocxMarkupMode) => {
								onMarkupChange(mode)
							}}
						>
							<DropdownMenuLabel>{t("previewDocxMarkup")}</DropdownMenuLabel>
							{DOCX_MARKUP_MODES.map(info => (
								<DropdownMenuRadioItem
									key={info.mode}
									value={info.mode}
									closeOnClick
								>
									{t(MARKUP_LABELS[info.mode])}
								</DropdownMenuRadioItem>
							))}
						</DropdownMenuRadioGroup>
					</>
				)}
			</DropdownMenuContent>
		</DropdownMenu>
	)
}
