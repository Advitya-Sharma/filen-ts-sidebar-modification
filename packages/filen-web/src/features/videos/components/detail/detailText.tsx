import { useState } from "react"
import { useTranslation } from "react-i18next"
import { cn } from "@filen/shared"
import { Button } from "@/components/ui/button"
import { Skeleton } from "@/components/ui/skeleton"

// The NFO's plot (clamped, with More) and cast as names; lines held for them while an NFO the index knows of is read,
// nothing at all once it has neither.
export function DetailSynopsis({ plot, cast, pending }: { plot: string | null; cast: readonly string[]; pending: boolean }) {
	const { t } = useTranslation("videos")
	const [expanded, setExpanded] = useState(false)

	if (pending) {
		return (
			<div
				aria-hidden="true"
				className="flex max-w-3xl flex-col gap-2"
			>
				<Skeleton className="h-4 w-full" />
				<Skeleton className="h-4 w-11/12" />
				<Skeleton className="h-4 w-2/3" />
				<Skeleton className="mt-1 h-4 w-1/3" />
			</div>
		)
	}

	if (plot === null && cast.length === 0) {
		return null
	}

	return (
		<div className="flex max-w-3xl flex-col gap-2 text-sm">
			{plot === null ? null : (
				<div className="flex flex-col items-start gap-1">
					<p className={cn("leading-6 whitespace-pre-line", !expanded && "line-clamp-4")}>{plot}</p>
					{plot.length > 280 ? (
						<Button
							variant="link"
							size="xs"
							className="px-0"
							aria-expanded={expanded}
							onClick={() => {
								setExpanded(!expanded)
							}}
						>
							{expanded ? t("detailLess") : t("detailMore")}
						</Button>
					) : null}
				</div>
			)}
			{cast.length === 0 ? null : <p className="text-muted-foreground">{t("detailCast", { names: cast.join(", ") })}</p>}
		</div>
	)
}
