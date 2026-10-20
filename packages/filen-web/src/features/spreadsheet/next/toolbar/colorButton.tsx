import { useState, type ReactNode } from "react"
import type { SheetColorInput, SheetRgb } from "@filen/office/core/sheet"
import { Button } from "@/components/ui/button"
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover"
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip"
import { DocxColorGrid } from "@/features/preview/components/docxFormatControls"
import { keepFocus } from "@/features/preview/components/docxEditor.logic"
import { STANDARD_SHEET_SWATCHES, swatchKeyOf, type SheetSwatch } from "@/features/spreadsheet/next/toolbar/toolbar.logic"

const NONE_KEY = "none"

// The colour picker the DOCX editor's toolbar uses (theme colours with their tints, standard colours, and none),
// over the workbook's theme. `onPick` gets the colour, or null for none (automatic text, no fill).
export function SheetColorGrid({
	theme,
	rgb,
	noneLabel,
	onPick
}: {
	theme: readonly (readonly SheetSwatch[])[]
	// The colour in effect (marked), undefined for none.
	rgb: SheetRgb | undefined
	noneLabel: string
	onPick: (color: SheetColorInput | null) => void
}) {
	const all = [...theme.flat(), ...STANDARD_SHEET_SWATCHES]

	return (
		<DocxColorGrid
			theme={theme}
			standard={STANDARD_SHEET_SWATCHES}
			chosen={rgb === undefined ? NONE_KEY : swatchKeyOf(rgb, all)}
			none={{ label: noneLabel, key: NONE_KEY }}
			onPick={swatch => {
				onPick(swatch === null ? null : (all.find(entry => entry.key === swatch.key)?.input ?? null))
			}}
		/>
	)
}

// A toolbar button opening the colour picker: its icon over a bar of the colour in effect.
export function SheetColorButton({
	label,
	icon,
	rgb,
	disabled,
	theme,
	noneLabel,
	onPick
}: {
	label: string
	icon: ReactNode
	rgb: SheetRgb | undefined
	disabled: boolean
	theme: readonly (readonly SheetSwatch[])[]
	noneLabel: string
	onPick: (color: SheetColorInput | null) => void
}) {
	const [open, setOpen] = useState(false)

	return (
		<Popover
			open={open}
			onOpenChange={setOpen}
		>
			<Tooltip>
				<TooltipTrigger
					render={
						<PopoverTrigger
							render={
								<Button
									variant="ghost"
									size="icon-sm"
									aria-label={label}
									disabled={disabled}
									className="shrink-0 flex-col gap-0"
									onMouseDown={keepFocus}
								>
									{icon}
									<span
										aria-hidden="true"
										className="h-0.5 w-3.5 rounded-full ring-1 ring-border"
										style={{
											backgroundColor: rgb === undefined ? "transparent" : `#${rgb.toString(16).padStart(6, "0")}`
										}}
									/>
								</Button>
							}
						/>
					}
				/>
				<TooltipContent>{label}</TooltipContent>
			</Tooltip>
			<PopoverContent
				data-preview-surface
				className="w-auto gap-2 p-3"
			>
				<SheetColorGrid
					theme={theme}
					rgb={rgb}
					noneLabel={noneLabel}
					onPick={color => {
						setOpen(false)
						onPick(color)
					}}
				/>
			</PopoverContent>
		</Popover>
	)
}
