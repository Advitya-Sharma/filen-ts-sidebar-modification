import { useEffect, useState, type ReactNode } from "react"
import { useTranslation } from "react-i18next"
import {
	BaselineIcon,
	BoldIcon,
	DecimalsArrowLeftIcon,
	DecimalsArrowRightIcon,
	ItalicIcon,
	MoreHorizontalIcon,
	PaintBucketIcon,
	PercentIcon,
	Redo2Icon,
	StrikethroughIcon,
	UnderlineIcon,
	Undo2Icon,
	WrapTextIcon,
	type LucideIcon
} from "lucide-react"
import type { SheetBorderEdges, SheetCommand, SheetFormatChoices, SheetFormatPatch, SheetRenderStyle } from "@filen/office/core/sheet"
import { entryText, roundedGeneral, type SheetGridModel, type SheetGridState } from "@filen/office/react/sheet"
import { Button } from "@/components/ui/button"
import {
	DropdownMenu,
	DropdownMenuCheckboxItem,
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
import { TooltipIconButton } from "@/components/ui/tooltipIconButton"
import { keepFocus } from "@/features/preview/components/docxEditor.logic"
import { FormatCellsDialog } from "@/features/spreadsheet/next/dialogs/formatCells"
import { setFormatCellsOpen, useFormatCellsOpen } from "@/features/spreadsheet/next/toolbar/formatCellsState"
import { FindButton } from "@/features/spreadsheet/next/find/findButton"
import { FreezeMenu } from "@/features/spreadsheet/next/menus/freezeMenu"
import { MergeMenu } from "@/features/spreadsheet/next/menus/mergeMenu"
import {
	HorizontalAlignItems,
	HorizontalAlignMenu,
	VerticalAlignItems,
	VerticalAlignMenu
} from "@/features/spreadsheet/next/toolbar/alignMenus"
import { BorderItems, BordersMenu, type BorderLine } from "@/features/spreadsheet/next/toolbar/bordersMenu"
import { SheetColorButton } from "@/features/spreadsheet/next/toolbar/colorButton"
import { FontPicker, SizePicker } from "@/features/spreadsheet/next/toolbar/fontControls"
import { NumberFormatItems, NumberFormatMenu } from "@/features/spreadsheet/next/toolbar/numberFormatMenu"
import {
	COMMA_PATCH,
	FONT_SIZES,
	PERCENT_PATCH,
	activeFormat,
	alignPatch,
	decimalsPatch,
	fontGroups,
	sampleOf,
	themeSwatchRows,
	toolbarFit,
	valignPatch,
	type ActiveFormat
} from "@/features/spreadsheet/next/toolbar/toolbar.logic"
import type { PreviewKey } from "@/lib/i18n"
import { useElementSize } from "@/lib/useElementSize"

// The spreadsheet toolbar: one row, as the DOCX editor's (undo and redo | font and size | bold, italic, underline,
// strikethrough | text colour, fill | borders | alignment and wrap | merge | number format and decimals | freeze |
// find). Groups the row has no room for move into the ⋯ menu, the last ones first; what still does not fit
// scrolls. The ⋯ menu also holds Comma Style, Format Cells and Clear Formats. Every control shows the active cell
// and applies to the selection.

interface ActiveCell {
	readonly format: ActiveFormat
	// The cell's number (its date as a serial), for the number format previews; null when it holds none.
	readonly sample: number | null
	// The point a General number shows, as its tile says ("." unless the session's locale has another).
	readonly decimal: string
}

// The active cell's format, asked of the session when the active cell or what it shows changes. The last cell's
// stays until this one's arrives, so the controls do not flicker while the arrows move on.
function useActiveCell(grid: SheetGridState, wanted: boolean): ActiveCell | undefined {
	const { model, sheet, selection, state } = grid
	const { row, col } = selection.active
	const [cell, setCell] = useState<ActiveCell | undefined>(undefined)
	const key = `${String(sheet)}:${String(row)}:${String(col)}:${String(state.version)}`

	useEffect(() => {
		if (!wanted) {
			return undefined
		}

		let live = true

		model.session.query("cell", { sheet, row, col }).then(
			detail => {
				if (!live) {
					return
				}

				const found = model.entry(sheet, row, col)
				const shown = found === null ? "" : entryText(found.tile, found.entry)
				const base: SheetRenderStyle = model.style(0)

				setCell({
					format: activeFormat(model.style(detail.style), base, detail.numFmt, shown),
					sample: sampleOf(detail.input, detail.flags, model.date1904),
					decimal: found?.tile.layers?.general?.decimal ?? "."
				})
			},
			() => {
				if (live) {
					setFormatCellsOpen(model, false)
				}
			}
		)

		return () => {
			live = false
		}
	}, [model, sheet, row, col, key, wanted])

	return wanted ? cell : undefined
}

// What the pickers offer from the workbook (its theme's colours and fonts, its own fonts): asked once.
function useFormatChoices(model: SheetGridModel, wanted: boolean): SheetFormatChoices | undefined {
	const [choices, setChoices] = useState<SheetFormatChoices | undefined>(undefined)

	useEffect(() => {
		if (!wanted) {
			return undefined
		}

		let live = true

		model.session.query("format.choices", {}).then(
			result => {
				if (live) {
					setChoices(result)
				}
			},
			() => undefined
		)

		return () => {
			live = false
		}
	}, [model, wanted])

	return choices
}

function Divider() {
	return (
		<span
			aria-hidden="true"
			className="mx-1 h-5 w-px shrink-0 bg-border"
		/>
	)
}

const TOGGLES: readonly {
	readonly key: "bold" | "italic" | "underline" | "strike"
	readonly icon: LucideIcon
	readonly label: PreviewKey
}[] = [
	{ key: "bold", icon: BoldIcon, label: "previewSpreadsheetBold" },
	{ key: "italic", icon: ItalicIcon, label: "previewSpreadsheetItalic" },
	{ key: "underline", icon: UnderlineIcon, label: "previewSpreadsheetUnderline" },
	{ key: "strike", icon: StrikethroughIcon, label: "previewSpreadsheetStrikethrough" }
]

function togglePatch(key: (typeof TOGGLES)[number]["key"], on: boolean): SheetFormatPatch {
	return key === "underline" ? { underline: on ? null : "single" } : { [key]: !on }
}

// Undo and redo for every editable sheet; the formats too for a workbook (a text file holds none). Shown disabled
// while editing waits to become possible, so the grid below does not move when it does.
export function SheetToolbar({
	grid,
	disabled,
	formats,
	onFormat
}: {
	grid: SheetGridState
	disabled: boolean
	formats: boolean
	onFormat: (patch: SheetFormatPatch) => void
}) {
	const { t } = useTranslation("preview")
	const model: SheetGridModel = grid.model
	const wanted = formats && !disabled
	const cell = useActiveCell(grid, wanted)
	const choices = useFormatChoices(model, formats)
	const [row, setRow] = useState<HTMLDivElement | null>(null)
	const { width } = useElementSize(row)
	const fit = toolbarFit(width)
	const [line, setLine] = useState<BorderLine>({ style: "thin", color: null, rgb: undefined })
	const dialogOpen = useFormatCellsOpen(model)
	const setDialogOpen = (open: boolean): void => {
		setFormatCellsOpen(model, open)
	}

	// Ctrl+1 while the toolbar is off has no cell format to start from: the dialog does not wait to open later.
	useEffect(() => {
		if (dialogOpen && !wanted) {
			setFormatCellsOpen(model, false)
		}
	}, [model, dialogOpen, wanted])
	const theme = choices === undefined ? [] : themeSwatchRows(choices.themeColors)
	const active = cell?.format
	const off = disabled || active === undefined
	const automatic = t("previewDocxColorAutomatic")
	const noFill = t("previewSpreadsheetColorNone")

	function command(next: SheetCommand): void {
		void model.intent()
		void model.command(next)
	}

	function borders(edges: SheetBorderEdges): void {
		command({ type: "format.borders", sheet: grid.sheet, ranges: grid.selection.ranges, edges })
	}

	function decimals(delta: 1 | -1): void {
		if (active === undefined) {
			return
		}

		// A General number a narrow cell rounds starts from what it shows, as Excel's does.
		const { row, col } = grid.selection.active
		const patch = decimalsPatch(active.numFmt, roundedGeneral(model, grid.sheet, row, col) ?? active.shown, delta, cell?.decimal ?? ".")

		if (patch !== null) {
			onFormat(patch)
		}
	}

	const history = (
		<>
			<TooltipIconButton
				label={t("previewSpreadsheetUndo")}
				disabled={disabled || !grid.state.canUndo}
				onMouseDown={keepFocus}
				onClick={() => {
					void model.undo()
				}}
			>
				<Undo2Icon />
			</TooltipIconButton>
			<TooltipIconButton
				label={t("previewSpreadsheetRedo")}
				disabled={disabled || !grid.state.canRedo}
				onMouseDown={keepFocus}
				onClick={() => {
					void model.redo()
				}}
			>
				<Redo2Icon />
			</TooltipIconButton>
		</>
	)

	if (!formats) {
		return (
			<div
				role="toolbar"
				aria-label={t("previewSheetFormatToolbar")}
				className="flex h-10 shrink-0 items-center gap-0.5 border-b border-border px-2"
			>
				{history}
			</div>
		)
	}

	const shown: ReactNode[] = []
	const group = (key: string, content: ReactNode): void => {
		shown.push(
			<Divider key={`${key}-divider`} />,
			<span
				key={key}
				className="contents"
			>
				{content}
			</span>
		)
	}

	if (fit.font) {
		group(
			"font",
			<>
				<FontPicker
					choices={choices}
					font={active?.font ?? ""}
					disabled={off}
					onFont={font => {
						onFormat({ font })
					}}
				/>
				<SizePicker
					size={active?.size ?? 11}
					disabled={off}
					onSize={size => {
						onFormat({ size })
					}}
				/>
			</>
		)
	}

	if (fit.toggles) {
		group(
			"toggles",
			TOGGLES.map(toggle => {
				const Icon = toggle.icon
				const on = active?.[toggle.key] === true

				return (
					<TooltipIconButton
						key={toggle.key}
						label={t(toggle.label)}
						disabled={off}
						pressed={on}
						onMouseDown={keepFocus}
						onClick={() => {
							onFormat(togglePatch(toggle.key, on))
						}}
					>
						<Icon />
					</TooltipIconButton>
				)
			})
		)
	}

	group(
		"colors",
		<>
			<SheetColorButton
				label={t("previewSpreadsheetTextColor")}
				icon={<BaselineIcon />}
				rgb={active?.color}
				disabled={off}
				theme={theme}
				noneLabel={automatic}
				onPick={color => {
					onFormat({ color })
				}}
			/>
			<SheetColorButton
				label={t("previewSpreadsheetFillColor")}
				icon={<PaintBucketIcon />}
				rgb={active?.fill}
				disabled={off}
				theme={theme}
				noneLabel={noFill}
				onPick={fill => {
					onFormat({ fill })
				}}
			/>
		</>
	)

	if (fit.borders) {
		group(
			"borders",
			<BordersMenu
				disabled={off}
				line={line}
				theme={theme}
				automaticLabel={automatic}
				onApply={borders}
				onLine={setLine}
			/>
		)
	}

	if (fit.align) {
		group(
			"align",
			<>
				<HorizontalAlignMenu
					align={active?.align ?? "general"}
					disabled={off}
					onAlign={align => {
						onFormat(alignPatch(align))
					}}
				/>
				<VerticalAlignMenu
					valign={active?.valign ?? "bottom"}
					disabled={off}
					onValign={valign => {
						onFormat(valignPatch(valign))
					}}
				/>
				<TooltipIconButton
					label={t("previewSheetFormatWrap")}
					disabled={off}
					pressed={active?.wrap === true}
					onMouseDown={keepFocus}
					onClick={() => {
						onFormat({ wrap: active?.wrap !== true })
					}}
				>
					<WrapTextIcon />
				</TooltipIconButton>
			</>
		)
	}

	group(
		"merge",
		<MergeMenu
			grid={grid}
			disabled={disabled}
		/>
	)

	if (fit.number) {
		group(
			"number",
			<>
				<NumberFormatMenu
					numFmt={active?.numFmt ?? "General"}
					sample={cell?.sample ?? null}
					date1904={model.date1904}
					disabled={off}
					onFormat={numFmt => {
						onFormat({ numFmt })
					}}
					onMore={() => {
						setDialogOpen(true)
					}}
				/>
				{fit.decimals ? (
					<>
						<TooltipIconButton
							label={t("previewSheetDecreaseDecimal")}
							disabled={off}
							onMouseDown={keepFocus}
							onClick={() => {
								decimals(-1)
							}}
						>
							<DecimalsArrowLeftIcon />
						</TooltipIconButton>
						<TooltipIconButton
							label={t("previewSheetIncreaseDecimal")}
							disabled={off}
							onMouseDown={keepFocus}
							onClick={() => {
								decimals(1)
							}}
						>
							<DecimalsArrowRightIcon />
						</TooltipIconButton>
						<TooltipIconButton
							label={t("previewSheetPercentStyle")}
							disabled={off}
							onMouseDown={keepFocus}
							onClick={() => {
								onFormat(PERCENT_PATCH)
							}}
						>
							<PercentIcon />
						</TooltipIconButton>
					</>
				) : null}
			</>
		)
	}

	group(
		"freeze",
		<FreezeMenu
			grid={grid}
			disabled={disabled}
		/>
	)

	return (
		<div
			ref={setRow}
			role="toolbar"
			aria-label={t("previewSheetFormatToolbar")}
			className="flex h-10 shrink-0 items-center gap-0.5 border-b border-border px-2"
		>
			<div className="-mx-1 flex min-w-0 flex-1 [scrollbar-width:none] items-center gap-0.5 self-stretch overflow-x-auto px-1 [&::-webkit-scrollbar]:hidden">
				{history}
				{shown}
				<FindButton
					grid={grid}
					disabled={disabled}
				/>
			</div>
			<DropdownMenu>
				<Tooltip>
					<TooltipTrigger
						render={
							<DropdownMenuTrigger
								render={
									<Button
										variant="ghost"
										size="icon-sm"
										disabled={disabled}
										aria-label={t("previewSheetFormatMore")}
										className="shrink-0"
									>
										<MoreHorizontalIcon />
									</Button>
								}
							/>
						}
					/>
					<TooltipContent>{t("previewSheetFormatMore")}</TooltipContent>
				</Tooltip>
				<DropdownMenuContent
					align="end"
					data-preview-surface
				>
					{fit.font ? null : (
						<>
							<DropdownMenuSub>
								<DropdownMenuSubTrigger disabled={off}>{t("previewSheetFormatFont")}</DropdownMenuSubTrigger>
								<DropdownMenuSubContent className="max-h-80 overflow-y-auto">
									<DropdownMenuRadioGroup
										value={active?.font ?? ""}
										onValueChange={(font: string) => {
											onFormat({ font })
										}}
									>
										{fontGroups(choices)
											.flatMap(entry => entry.fonts)
											.map(font => (
												<DropdownMenuRadioItem
													key={font.name}
													value={font.name}
													closeOnClick
												>
													{font.name}
												</DropdownMenuRadioItem>
											))}
									</DropdownMenuRadioGroup>
								</DropdownMenuSubContent>
							</DropdownMenuSub>
							<DropdownMenuSub>
								<DropdownMenuSubTrigger disabled={off}>{t("previewSheetFormatFontSize")}</DropdownMenuSubTrigger>
								<DropdownMenuSubContent className="max-h-80 overflow-y-auto">
									<DropdownMenuRadioGroup
										value={String(active?.size ?? "")}
										onValueChange={(size: string) => {
											onFormat({ size: Number(size) })
										}}
									>
										{FONT_SIZES.map(size => (
											<DropdownMenuRadioItem
												key={size}
												value={String(size)}
												closeOnClick
											>
												{size}
											</DropdownMenuRadioItem>
										))}
									</DropdownMenuRadioGroup>
								</DropdownMenuSubContent>
							</DropdownMenuSub>
						</>
					)}
					{fit.toggles
						? null
						: TOGGLES.map(toggle => {
								const Icon = toggle.icon
								const on = active?.[toggle.key] === true

								return (
									<DropdownMenuCheckboxItem
										key={toggle.key}
										checked={on}
										closeOnClick
										disabled={off}
										onCheckedChange={() => {
											onFormat(togglePatch(toggle.key, on))
										}}
									>
										<Icon />
										{t(toggle.label)}
									</DropdownMenuCheckboxItem>
								)
							})}
					{fit.borders ? null : (
						<DropdownMenuSub>
							<DropdownMenuSubTrigger disabled={off}>{t("previewSheetFormatBorders")}</DropdownMenuSubTrigger>
							<DropdownMenuSubContent>
								<BorderItems
									line={line}
									theme={theme}
									automaticLabel={automatic}
									onApply={borders}
									onLine={setLine}
								/>
							</DropdownMenuSubContent>
						</DropdownMenuSub>
					)}
					{fit.align ? null : (
						<>
							<DropdownMenuSub>
								<DropdownMenuSubTrigger disabled={off}>{t("previewSheetFormatHorizontalAlign")}</DropdownMenuSubTrigger>
								<DropdownMenuSubContent>
									<HorizontalAlignItems
										align={active?.align ?? "general"}
										onAlign={align => {
											onFormat(alignPatch(align))
										}}
									/>
								</DropdownMenuSubContent>
							</DropdownMenuSub>
							<DropdownMenuSub>
								<DropdownMenuSubTrigger disabled={off}>{t("previewSheetFormatVerticalAlign")}</DropdownMenuSubTrigger>
								<DropdownMenuSubContent>
									<VerticalAlignItems
										valign={active?.valign ?? "bottom"}
										onValign={valign => {
											onFormat(valignPatch(valign))
										}}
									/>
								</DropdownMenuSubContent>
							</DropdownMenuSub>
							<DropdownMenuCheckboxItem
								checked={active?.wrap === true}
								closeOnClick
								disabled={off}
								onCheckedChange={() => {
									onFormat({ wrap: active?.wrap !== true })
								}}
							>
								<WrapTextIcon />
								{t("previewSheetFormatWrap")}
							</DropdownMenuCheckboxItem>
						</>
					)}
					{fit.number ? null : (
						<DropdownMenuSub>
							<DropdownMenuSubTrigger disabled={off}>{t("previewSpreadsheetNumberFormat")}</DropdownMenuSubTrigger>
							<DropdownMenuSubContent>
								<NumberFormatItems
									numFmt={active?.numFmt ?? "General"}
									sample={cell?.sample ?? null}
									date1904={model.date1904}
									onFormat={numFmt => {
										onFormat({ numFmt })
									}}
									onMore={() => {
										setDialogOpen(true)
									}}
								/>
							</DropdownMenuSubContent>
						</DropdownMenuSub>
					)}
					{fit.decimals && fit.number ? null : (
						<>
							<DropdownMenuItem
								disabled={off}
								onClick={() => {
									decimals(-1)
								}}
							>
								<DecimalsArrowLeftIcon />
								{t("previewSheetDecreaseDecimal")}
							</DropdownMenuItem>
							<DropdownMenuItem
								disabled={off}
								onClick={() => {
									decimals(1)
								}}
							>
								<DecimalsArrowRightIcon />
								{t("previewSheetIncreaseDecimal")}
							</DropdownMenuItem>
							<DropdownMenuItem
								disabled={off}
								onClick={() => {
									onFormat(PERCENT_PATCH)
								}}
							>
								<PercentIcon />
								{t("previewSheetPercentStyle")}
							</DropdownMenuItem>
						</>
					)}
					{fit.font && fit.toggles && fit.borders && fit.align && fit.number && fit.decimals ? null : <DropdownMenuSeparator />}
					<DropdownMenuItem
						disabled={off}
						onClick={() => {
							onFormat(COMMA_PATCH)
						}}
					>
						{t("previewSheetCommaStyle")}
					</DropdownMenuItem>
					<DropdownMenuItem
						disabled={off}
						onClick={() => {
							setDialogOpen(true)
						}}
					>
						{t("previewSheetFormatCells")}
					</DropdownMenuItem>
					<DropdownMenuItem
						disabled={disabled}
						onClick={() => {
							command({ type: "format.clear", sheet: grid.sheet, ranges: grid.selection.ranges })
						}}
					>
						{t("previewSheetClearFormats")}
					</DropdownMenuItem>
				</DropdownMenuContent>
			</DropdownMenu>
			<FormatCellsDialog
				open={dialogOpen && active !== undefined}
				numFmt={active?.numFmt ?? "General"}
				sample={cell?.sample ?? null}
				date1904={model.date1904}
				onOpenChange={setDialogOpen}
				onApply={numFmt => {
					onFormat({ numFmt })
				}}
			/>
		</div>
	)
}
