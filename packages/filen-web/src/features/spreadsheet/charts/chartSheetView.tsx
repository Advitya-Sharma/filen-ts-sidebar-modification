import { useState } from "react"
import { useTranslation } from "react-i18next"
import { SheetChart, type SheetChartSource } from "@filen/office/react/sheet/charts"
import { useElementSize } from "@/lib/useElementSize"

// A chartsheet: its chart over the whole sheet area, at the size its drawing gives it, scaled to fit as Excel's "Zoom
// with window" shows it. A double-click, or Enter once it has focus, opens the chart editor.

const EMU_PER_PX = 9525

export interface ChartSheetViewProps {
	readonly part: string
	// EMU.
	readonly width: number
	readonly height: number
	readonly source: SheetChartSource
	readonly version: number
	readonly onEdit: () => void
}

export function ChartSheetView(props: ChartSheetViewProps) {
	const { t } = useTranslation("preview")
	const [host, setHost] = useState<HTMLDivElement | null>(null)
	const size = useElementSize(host)
	const width = props.width / EMU_PER_PX
	const height = props.height / EMU_PER_PX
	const scale = width > 0 && height > 0 ? Math.min(size.width / width, size.height / height) : 0

	return (
		<div
			ref={setHost}
			className="pointer-events-auto absolute inset-0 flex items-center justify-center bg-background p-4"
		>
			{scale > 0 ? (
				<div
					role="button"
					tabIndex={0}
					aria-label={t("previewSheetChartEdit")}
					className="relative shrink-0 rounded-sm outline-none focus-visible:ring-3 focus-visible:ring-ring/50"
					style={{ width: width * scale, height: height * scale }}
					onDoubleClick={props.onEdit}
					onKeyDown={event => {
						if (event.key !== "Enter" && event.key !== " ") return

						event.preventDefault()
						event.stopPropagation()
						props.onEdit()
					}}
				>
					<SheetChart
						part={props.part}
						width={width}
						height={height}
						source={props.source}
						version={props.version}
					/>
				</div>
			) : null}
		</div>
	)
}
