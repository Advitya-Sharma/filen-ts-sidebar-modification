import { useTranslation } from "react-i18next"
import type { SheetFormatChoices } from "@filen/office/core/sheet"
import { Select, SelectContent, SelectGroup, SelectItem, SelectLabel, SelectTrigger, SelectValue } from "@/components/ui/select"
import { i18n } from "@/lib/i18n"
import { FONT_SIZES, fontGroups, type FontGroup } from "@/features/spreadsheet/next/toolbar/toolbar.logic"

// The font and size pickers, as the DOCX editor's toolbar has them. A font the lists lack (the cell's own) shows as
// an item of its own.

function fontLabel(font: FontGroup["fonts"][number]): string {
	if (font.theme === "headings") {
		return i18n.t("preview:previewSheetFormatHeadingsFont", { font: font.name })
	}

	if (font.theme === "body") {
		return i18n.t("preview:previewSheetFormatBodyFont", { font: font.name })
	}

	return font.name
}

export function FontPicker({
	choices,
	font,
	disabled,
	onFont
}: {
	choices: SheetFormatChoices | undefined
	font: string
	disabled: boolean
	onFont: (font: string) => void
}) {
	const { t } = useTranslation("preview")
	const groups = fontGroups(choices)
	const listed = groups.flatMap(group => group.fonts)
	const own = listed.some(entry => entry.name.toLowerCase() === font.toLowerCase()) ? null : font
	const items = [
		...listed.map(entry => ({ value: entry.name, label: fontLabel(entry) })),
		...(own === null ? [] : [{ value: own, label: own }])
	]
	const value = listed.find(entry => entry.name.toLowerCase() === font.toLowerCase())?.name ?? font

	return (
		<Select
			items={items}
			value={value}
			disabled={disabled}
			onValueChange={next => {
				if (typeof next === "string" && next !== value) {
					onFont(next)
				}
			}}
		>
			<SelectTrigger
				size="sm"
				aria-label={t("previewSheetFormatFont")}
				className="w-36 shrink-0"
			>
				<SelectValue />
			</SelectTrigger>
			<SelectContent
				data-preview-surface
				aria-label={t("previewSheetFormatFont")}
				className="max-h-80 min-w-48"
			>
				{groups.map(group => (
					<SelectGroup key={group.label}>
						<SelectLabel>{t(group.label)}</SelectLabel>
						{group.fonts.map(entry => (
							<SelectItem
								key={entry.name}
								value={entry.name}
							>
								{fontLabel(entry)}
							</SelectItem>
						))}
					</SelectGroup>
				))}
				{own === null ? null : (
					<SelectGroup>
						<SelectItem value={own}>{own}</SelectItem>
					</SelectGroup>
				)}
			</SelectContent>
		</Select>
	)
}

export function SizePicker({ size, disabled, onSize }: { size: number; disabled: boolean; onSize: (size: number) => void }) {
	const { t } = useTranslation("preview")
	const sizes = FONT_SIZES.includes(size) ? FONT_SIZES : [...FONT_SIZES, size].sort((a, b) => a - b)
	const items = sizes.map(entry => ({ value: String(entry), label: i18n.format(entry, "number") }))

	return (
		<Select
			items={items}
			value={String(size)}
			disabled={disabled}
			onValueChange={next => {
				if (typeof next === "string" && Number(next) !== size) {
					onSize(Number(next))
				}
			}}
		>
			<SelectTrigger
				size="sm"
				aria-label={t("previewSheetFormatFontSize")}
				className="w-16 shrink-0"
			>
				<SelectValue />
			</SelectTrigger>
			<SelectContent
				aria-label={t("previewSheetFormatFontSize")}
				data-preview-surface
				className="max-h-80"
			>
				<SelectGroup>
					{sizes.map(entry => (
						<SelectItem
							key={entry}
							value={String(entry)}
						>
							{i18n.format(entry, "number")}
						</SelectItem>
					))}
				</SelectGroup>
			</SelectContent>
		</Select>
	)
}
