import { useEffect } from "react"
import { useTranslation } from "react-i18next"
import { CaptionsIcon, CropIcon, RotateCwIcon, Trash2Icon, WrapTextIcon } from "lucide-react"
import type { DocxEditorCommandType } from "@filen/office/core/docx/edit"
import { cn } from "@filen/shared"
import { keepFocus, type DocxDrawingItem, type DocxToolbarSlotProps } from "@/features/preview/components/docxEditor.logic"
import { ROTATE_VARIANTS, WRAP_VARIANTS } from "@/features/preview/components/docxCommandLabels/drawings"
import { setDrawingCrop, useDrawingCrop } from "@/features/preview/components/docxDrawingCrop"
import { useCommandMenuFocus } from "@/features/preview/components/docxInsert.logic"
import { Button } from "@/components/ui/button"
import {
	DropdownMenu,
	DropdownMenuCheckboxItem,
	DropdownMenuContent,
	DropdownMenuRadioGroup,
	DropdownMenuRadioItem,
	DropdownMenuItem,
	DropdownMenuSeparator,
	DropdownMenuSub,
	DropdownMenuSubContent,
	DropdownMenuSubTrigger,
	DropdownMenuTrigger
} from "@/components/ui/dropdown-menu"
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip"
import { TooltipIconButton } from "@/components/ui/tooltipIconButton"

// The toolbar's contextual Picture group while a picture is selected: Wrap text, Rotate, Crop (page view: its
// handles on the page), Alt text and Delete; resizing and moving are the handles'. In Mobile view the group
// shows without cropping. The context menu asks for crop and alt text through the input channel too. On a
// narrow row the last items move to the ⋯ menu (DocxDrawingOverflow), as the toolbar's other groups do.

export interface DocxDrawingMenuProps extends DocxToolbarSlotProps {
	// The items shown here (drawingItems, as many as the row has room for).
	items: readonly DocxDrawingItem[]
}

export function DocxDrawingMenu(props: DocxDrawingMenuProps) {
	const { t } = useTranslation("preview")
	const { commands, snapshot, readOnly, execute, requestInput, inputRequest, onInputDone, surface, items } = props
	const drawing = snapshot?.context.drawing ?? null
	const pos = drawing?.pos ?? null
	const crop = useDrawingCrop(execute)
	const focus = useCommandMenuFocus()
	const has = (type: DocxEditorCommandType): boolean => commands.some(command => command.info.type === type)
	const off = (type: DocxEditorCommandType): boolean => snapshot === undefined || readOnly || !snapshot.enabled[type]
	const canCrop = surface === "pages" && has("drawing.crop") && drawing?.croppable === true && !off("drawing.crop")
	const cropping = pos !== null && crop === pos
	const cropRequest = inputRequest?.type === "drawing.crop" ? inputRequest.nonce : null

	// Crop mode ends with the selection on another picture, or none, or once cropping is not possible there.
	useEffect(() => {
		if (!canCrop) {
			setDrawingCrop(execute, null)
		}

		return () => {
			setDrawingCrop(execute, null)
		}
	}, [pos, canCrop, execute])

	// The context menu's Crop.
	useEffect(() => {
		if (cropRequest === null) {
			return
		}

		onInputDone(cropRequest)

		if (canCrop && pos !== null) {
			setDrawingCrop(execute, pos)
		}
	}, [cropRequest, canCrop, pos, onInputDone, execute])

	// Enter or Escape leaves crop mode, as in Word.
	useEffect(() => {
		if (!cropping) {
			return
		}

		// Keys pressed in the document only (its pages and their input): the find field, a toolbar button, a dialog
		// or a menu over it keep their own Enter and Escape.
		const onKey = (event: KeyboardEvent): void => {
			if (!(event.target instanceof Element) || event.target.closest(".fo-docx-editor") === null) {
				return
			}

			if (event.key === "Escape" || event.key === "Enter") {
				event.preventDefault()
				event.stopPropagation()
				setDrawingCrop(execute, null)
			}
		}

		window.addEventListener("keydown", onKey, true)

		return () => {
			window.removeEventListener("keydown", onKey, true)
		}
	}, [cropping, execute])

	if (drawing === null || items.length === 0) {
		return null
	}

	return (
		<div
			role="group"
			aria-label={t("previewDocxPictureGroup")}
			data-docx-picture-group
			className="flex shrink-0 items-center gap-0.5"
		>
			<span
				aria-hidden="true"
				className="mx-1 h-5 w-px shrink-0 bg-border"
			/>
			{items.includes("drawing.wrap") ? (
				<DropdownMenu>
					<Tooltip>
						<TooltipTrigger
							render={
								<DropdownMenuTrigger
									render={
										<Button
											variant="ghost"
											size="icon-sm"
											aria-label={t("previewDocxPictureWrap")}
											disabled={off("drawing.wrap")}
											onMouseDown={keepFocus}
										>
											<WrapTextIcon />
										</Button>
									}
								/>
							}
						/>
						<TooltipContent>{t("previewDocxPictureWrap")}</TooltipContent>
					</Tooltip>
					<DropdownMenuContent
						data-preview-surface
						finalFocus={focus.finalFocus}
					>
						<DropdownMenuRadioGroup value={drawing.wrap}>
							{WRAP_VARIANTS.map(variant =>
								"command" in variant && variant.command.type === "drawing.wrap" ? (
									<DropdownMenuRadioItem
										key={variant.label}
										value={variant.command.mode}
										closeOnClick
										onClick={() => {
											execute(variant.command)
											focus.ran()
										}}
									>
										{t(variant.label)}
									</DropdownMenuRadioItem>
								) : null
							)}
						</DropdownMenuRadioGroup>
					</DropdownMenuContent>
				</DropdownMenu>
			) : null}
			{items.includes("drawing.rotate") ? (
				<DropdownMenu>
					<Tooltip>
						<TooltipTrigger
							render={
								<DropdownMenuTrigger
									render={
										<Button
											variant="ghost"
											size="icon-sm"
											aria-label={t("previewDocxPictureRotate")}
											disabled={off("drawing.rotate")}
											onMouseDown={keepFocus}
										>
											<RotateCwIcon />
										</Button>
									}
								/>
							}
						/>
						<TooltipContent>{t("previewDocxPictureRotate")}</TooltipContent>
					</Tooltip>
					<DropdownMenuContent
						data-preview-surface
						finalFocus={focus.finalFocus}
					>
						{ROTATE_VARIANTS.map(variant =>
							"command" in variant ? (
								<DropdownMenuItem
									key={variant.label}
									onClick={() => {
										execute(variant.command)
										focus.ran()
									}}
								>
									{t(variant.label)}
								</DropdownMenuItem>
							) : null
						)}
					</DropdownMenuContent>
				</DropdownMenu>
			) : null}
			{items.includes("drawing.crop") ? (
				<TooltipIconButton
					label={t("previewDocxPictureCrop")}
					disabled={!canCrop}
					pressed={cropping}
					className={cn(cropping && "bg-accent text-foreground")}
					onMouseDown={keepFocus}
					onClick={() => {
						setDrawingCrop(execute, cropping ? null : drawing.pos)
					}}
				>
					<CropIcon />
				</TooltipIconButton>
			) : null}
			{items.includes("drawing.altText") ? (
				<TooltipIconButton
					label={t("previewDocxPictureAltText")}
					disabled={off("drawing.altText")}
					onMouseDown={keepFocus}
					onClick={() => {
						requestInput("drawing.altText")
					}}
				>
					<CaptionsIcon />
				</TooltipIconButton>
			) : null}
			{items.includes("drawing.delete") ? (
				<TooltipIconButton
					label={t("previewDocxPictureDelete")}
					disabled={off("drawing.delete")}
					onMouseDown={keepFocus}
					onClick={() => {
						execute({ type: "drawing.delete" })
					}}
				>
					<Trash2Icon />
				</TooltipIconButton>
			) : null}
		</div>
	)
}

// The Picture group's items the row has no room for, in the ⋯ menu (`items`), with a separator after them.
export function DocxDrawingOverflow(props: DocxDrawingMenuProps) {
	const { t } = useTranslation("preview")
	const { snapshot, readOnly, execute, requestInput, items } = props
	const drawing = snapshot?.context.drawing ?? null
	const crop = useDrawingCrop(execute)

	if (drawing === null || items.length === 0) {
		return null
	}

	const off = (type: DocxEditorCommandType): boolean => snapshot === undefined || readOnly || !snapshot.enabled[type]
	const cropping = crop === drawing.pos

	return (
		<>
			{items.includes("drawing.wrap") ? (
				<DropdownMenuSub>
					<DropdownMenuSubTrigger disabled={off("drawing.wrap")}>{t("previewDocxPictureWrap")}</DropdownMenuSubTrigger>
					<DropdownMenuSubContent>
						<DropdownMenuRadioGroup value={drawing.wrap}>
							{WRAP_VARIANTS.map(variant =>
								"command" in variant && variant.command.type === "drawing.wrap" ? (
									<DropdownMenuRadioItem
										key={variant.label}
										value={variant.command.mode}
										closeOnClick
										onClick={() => {
											execute(variant.command)
										}}
									>
										{t(variant.label)}
									</DropdownMenuRadioItem>
								) : null
							)}
						</DropdownMenuRadioGroup>
					</DropdownMenuSubContent>
				</DropdownMenuSub>
			) : null}
			{items.includes("drawing.rotate") ? (
				<DropdownMenuSub>
					<DropdownMenuSubTrigger disabled={off("drawing.rotate")}>{t("previewDocxPictureRotate")}</DropdownMenuSubTrigger>
					<DropdownMenuSubContent>
						{ROTATE_VARIANTS.map(variant =>
							"command" in variant ? (
								<DropdownMenuItem
									key={variant.label}
									onClick={() => {
										execute(variant.command)
									}}
								>
									{t(variant.label)}
								</DropdownMenuItem>
							) : null
						)}
					</DropdownMenuSubContent>
				</DropdownMenuSub>
			) : null}
			{items.includes("drawing.crop") ? (
				<DropdownMenuCheckboxItem
					checked={cropping}
					closeOnClick
					disabled={!drawing.croppable || off("drawing.crop")}
					onCheckedChange={() => {
						setDrawingCrop(execute, cropping ? null : drawing.pos)
					}}
				>
					{t("previewDocxPictureCrop")}
				</DropdownMenuCheckboxItem>
			) : null}
			{items.includes("drawing.altText") ? (
				<DropdownMenuItem
					disabled={off("drawing.altText")}
					onClick={() => {
						requestInput("drawing.altText")
					}}
				>
					{t("previewDocxPictureAltText")}
				</DropdownMenuItem>
			) : null}
			{items.includes("drawing.delete") ? (
				<DropdownMenuItem
					disabled={off("drawing.delete")}
					onClick={() => {
						execute({ type: "drawing.delete" })
					}}
				>
					{t("previewDocxPictureDelete")}
				</DropdownMenuItem>
			) : null}
			<DropdownMenuSeparator />
		</>
	)
}
