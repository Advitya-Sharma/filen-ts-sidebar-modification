import { useTranslation } from "react-i18next"
import { ChevronDownIcon } from "lucide-react"
import {
	SHEET_BORDER_PRESETS,
	borderPresetEdges,
	type SheetBorderEdges,
	type SheetBorderLineInput,
	type SheetBorderPreset,
	type SheetBorderStyle,
	type SheetColorInput,
	type SheetRgb
} from "@filen/office/core/sheet"
import { Button } from "@/components/ui/button"
import {
	DropdownMenu,
	DropdownMenuContent,
	DropdownMenuItem,
	DropdownMenuRadioGroup,
	DropdownMenuRadioItem,
	DropdownMenuSeparator,
	DropdownMenuSub,
	DropdownMenuSubContent,
	DropdownMenuSubTrigger,
	DropdownMenuTrigger
} from "@/components/ui/dropdown-menu"
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip"
import type { PreviewKey } from "@/lib/i18n"
import { SheetColorGrid } from "@/features/spreadsheet/next/toolbar/colorButton"
import { BORDER_LABELS, LINE_STYLES, type SheetSwatch } from "@/features/spreadsheet/next/toolbar/toolbar.logic"

// The Borders menu: Excel's presets drawn as small grids, and the line colour and style the presets draw with.

// A 2×2 grid, faint, with the preset's edges drawn solid.
export function BorderIcon({ edges }: { edges: SheetBorderEdges }) {
	const on = (line: SheetBorderLineInput | null | undefined): boolean => line !== undefined && line !== null
	const width = (line: SheetBorderLineInput | null | undefined): number =>
		line?.style === "thick" ? 2.5 : line?.style === "double" || line?.style === "medium" ? 1.75 : 1.25

	return (
		<svg
			aria-hidden="true"
			viewBox="0 0 16 16"
			className="size-4 shrink-0"
			fill="none"
			stroke="currentColor"
		>
			<g
				strokeOpacity={0.3}
				strokeDasharray="1 1"
				strokeWidth={1}
			>
				<rect
					x={1.5}
					y={1.5}
					width={13}
					height={13}
				/>
				<path d="M8 1.5v13M1.5 8h13" />
			</g>
			{on(edges.top) ? (
				<path
					d="M1.5 1.5h13"
					strokeWidth={width(edges.top)}
				/>
			) : null}
			{on(edges.bottom) ? (
				<path
					d="M1.5 14.5h13"
					strokeWidth={width(edges.bottom)}
				/>
			) : null}
			{on(edges.left) ? (
				<path
					d="M1.5 1.5v13"
					strokeWidth={width(edges.left)}
				/>
			) : null}
			{on(edges.right) ? (
				<path
					d="M14.5 1.5v13"
					strokeWidth={width(edges.right)}
				/>
			) : null}
			{on(edges.insideHorizontal) ? (
				<path
					d="M1.5 8h13"
					strokeWidth={1.25}
				/>
			) : null}
			{on(edges.insideVertical) ? (
				<path
					d="M8 1.5v13"
					strokeWidth={1.25}
				/>
			) : null}
			{on(edges.diagonalUp) ? (
				<path
					d="M1.5 14.5l13-13"
					strokeWidth={1.25}
				/>
			) : null}
			{on(edges.diagonalDown) ? (
				<path
					d="M1.5 1.5l13 13"
					strokeWidth={1.25}
				/>
			) : null}
		</svg>
	)
}

// A line style as a sample stroke.
export function LineSample({ style }: { style: SheetBorderStyle }) {
	const line = LINE_STYLES.find(entry => entry.style === style)

	if (line === undefined) {
		return null
	}

	return (
		<svg
			aria-hidden="true"
			viewBox="0 0 48 8"
			className="h-2 w-12 shrink-0"
			stroke="currentColor"
		>
			{line.double === true ? (
				<path
					d="M0 2.5h48M0 5.5h48"
					strokeWidth={1}
				/>
			) : (
				<path
					d="M0 4h48"
					strokeWidth={line.width}
					strokeDasharray={line.dash === "" ? undefined : line.dash}
				/>
			)}
		</svg>
	)
}

// Presets in groups, as Excel's menu separates them.
const GROUPS: readonly { readonly label: PreviewKey; readonly presets: readonly SheetBorderPreset[] }[] = [
	{ label: "previewSheetBorderGroupEdges", presets: SHEET_BORDER_PRESETS.slice(0, 4) },
	{ label: "previewSheetBorderGroupOutlines", presets: SHEET_BORDER_PRESETS.slice(4, 8) },
	{ label: "previewSheetBorderGroupTopBottom", presets: SHEET_BORDER_PRESETS.slice(8, 13) },
	{ label: "previewSheetBorderGroupInside", presets: SHEET_BORDER_PRESETS.slice(13) }
]

export interface BorderLine {
	readonly style: SheetBorderStyle
	// null: automatic.
	readonly color: SheetColorInput | null
	// What the colour shows, for the picker's mark.
	readonly rgb: SheetRgb | undefined
}

// The menu's items, for the toolbar's dropdown and the ⋯ menu's submenu alike.
export function BorderItems({
	line,
	theme,
	automaticLabel,
	onApply,
	onLine
}: {
	line: BorderLine
	theme: readonly (readonly SheetSwatch[])[]
	automaticLabel: string
	onApply: (edges: SheetBorderEdges) => void
	onLine: (line: BorderLine) => void
}) {
	const { t } = useTranslation("preview")
	const input: SheetBorderLineInput = { style: line.style, color: line.color }

	return (
		<>
			{GROUPS.map((group, index) => (
				<div
					key={group.label}
					role="group"
					aria-label={t(group.label)}
				>
					{index > 0 ? <DropdownMenuSeparator /> : null}
					{group.presets.map(preset => {
						const edges = borderPresetEdges(preset, input)

						return (
							<DropdownMenuItem
								key={preset}
								onClick={() => {
									onApply(edges)
								}}
							>
								<BorderIcon edges={edges} />
								{t(BORDER_LABELS[preset])}
							</DropdownMenuItem>
						)
					})}
				</div>
			))}
			<DropdownMenuSeparator />
			<DropdownMenuSub>
				<DropdownMenuSubTrigger>{t("previewSheetBorderLineColor")}</DropdownMenuSubTrigger>
				<DropdownMenuSubContent className="w-auto p-3">
					<SheetColorGrid
						theme={theme}
						rgb={line.rgb}
						noneLabel={automaticLabel}
						onPick={color => {
							onLine({ ...line, color, rgb: color === null ? undefined : rgbOf(color, theme) })
						}}
					/>
				</DropdownMenuSubContent>
			</DropdownMenuSub>
			<DropdownMenuSub>
				<DropdownMenuSubTrigger>{t("previewSheetBorderLineStyle")}</DropdownMenuSubTrigger>
				<DropdownMenuSubContent>
					<DropdownMenuRadioGroup
						value={line.style}
						onValueChange={(style: SheetBorderStyle) => {
							onLine({ ...line, style })
						}}
					>
						{LINE_STYLES.map(entry => (
							<DropdownMenuRadioItem
								key={entry.style}
								value={entry.style}
								closeOnClick
								aria-label={t(entry.label)}
							>
								<LineSample style={entry.style} />
								{t(entry.label)}
							</DropdownMenuRadioItem>
						))}
					</DropdownMenuRadioGroup>
				</DropdownMenuSubContent>
			</DropdownMenuSub>
		</>
	)
}

// What a picked colour shows: a theme swatch's colour, or the colour itself.
function rgbOf(color: SheetColorInput, theme: readonly (readonly SheetSwatch[])[]): SheetRgb | undefined {
	if (typeof color === "number") {
		return color
	}

	const swatch = theme
		.flat()
		.find(
			entry => typeof entry.input !== "number" && entry.input.theme === color.theme && (entry.input.tint ?? 0) === (color.tint ?? 0)
		)

	return swatch === undefined ? undefined : Number.parseInt(swatch.rgb, 16)
}

export function BordersMenu({
	disabled,
	line,
	theme,
	automaticLabel,
	onApply,
	onLine
}: {
	disabled: boolean
	line: BorderLine
	theme: readonly (readonly SheetSwatch[])[]
	automaticLabel: string
	onApply: (edges: SheetBorderEdges) => void
	onLine: (line: BorderLine) => void
}) {
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
									aria-label={t("previewSheetFormatBorders")}
									className="shrink-0 gap-0.5 px-1.5"
								>
									<BorderIcon edges={borderPresetEdges("all")} />
									<ChevronDownIcon className="size-3" />
								</Button>
							}
						/>
					}
				/>
				<TooltipContent>{t("previewSheetFormatBorders")}</TooltipContent>
			</Tooltip>
			<DropdownMenuContent data-preview-surface>
				<BorderItems
					line={line}
					theme={theme}
					automaticLabel={automaticLabel}
					onApply={onApply}
					onLine={onLine}
				/>
			</DropdownMenuContent>
		</DropdownMenu>
	)
}
