import { useRef, useState, type ReactNode, type RefObject } from "react"
import { useTranslation } from "react-i18next"
import type { SheetValidationAlert } from "@filen/office/core/sheet"
import {
	AlertDialog,
	AlertDialogContent,
	AlertDialogDescription,
	AlertDialogFooter,
	AlertDialogHeader,
	AlertDialogTitle
} from "@/components/ui/alert-dialog"
import type { SheetGridHandle } from "@filen/office/react/sheet"
import { Button } from "@/components/ui/button"
import type { PreviewKey } from "@/lib/i18n"

// The alert an entry that breaks its cell's validation shows, with Excel's buttons for each style: Stop asks to
// retry (the entry reopens with what was typed) or cancel (it is dropped); Warning asks "Continue?" (Yes writes it,
// No reopens it, Cancel drops it); Information tells and writes it on OK. Closing it drops the entry, as Cancel.
// Entries checked while one alert is open wait their turn, each asked in the order it was made, and none interrupts an
// entry being typed (it shows once that entry is committed or cancelled); once none is left, typing goes on in the
// grid (or in the entry reopened).

export type ValidationVerdict = "commit" | "edit" | "discard"

export const ALERT_BUTTONS: Readonly<
	Record<SheetValidationAlert["style"], readonly { readonly key: PreviewKey | "cancel"; readonly verdict: ValidationVerdict }[]>
> = {
	stop: [
		{ key: "previewSheetValidationRetry", verdict: "edit" },
		{ key: "cancel", verdict: "discard" }
	],
	warning: [
		{ key: "previewSheetValidationYes", verdict: "commit" },
		{ key: "previewSheetValidationNo", verdict: "edit" },
		{ key: "cancel", verdict: "discard" }
	],
	information: [
		{ key: "previewSheetValidationOk", verdict: "commit" },
		{ key: "cancel", verdict: "discard" }
	]
}

interface Asked {
	readonly alert: SheetValidationAlert
	readonly resolve: (verdict: ValidationVerdict) => void
}

// An alert asked from anywhere (the grid's entry check), answered by the user: `ask` resolves with the verdict. Asks
// made while one is open are queued; none shows while `editing` (an entry is open in the grid).
export function useValidationAlert(
	gridRef: RefObject<SheetGridHandle | null>,
	editing: boolean
): {
	readonly ask: (alert: SheetValidationAlert) => Promise<ValidationVerdict>
	readonly dialog: ReactNode
} {
	const { t } = useTranslation(["preview", "common"])
	const [queue, setQueue] = useState<readonly Asked[]>([])
	// The last answer, for where focus goes once the dialog closes.
	const lastVerdict = useRef<ValidationVerdict>("discard")
	const asked = queue[0] ?? null

	function ask(alert: SheetValidationAlert): Promise<ValidationVerdict> {
		return new Promise(resolve => {
			setQueue(current => [...current, { alert, resolve }])
		})
	}

	function answer(verdict: ValidationVerdict): void {
		if (asked === null) return

		lastVerdict.current = verdict
		setQueue(current => current.filter(item => item !== asked))
		asked.resolve(verdict)
	}

	const alert = asked?.alert
	const dialog = (
		<AlertDialog
			open={asked !== null && !editing}
			onOpenChange={open => {
				if (!open) answer("discard")
			}}
		>
			<AlertDialogContent
				// The entry it asked about is gone: typing goes on in the grid, or in the entry Retry reopened (which
				// takes focus itself).
				finalFocus={() => {
					if (lastVerdict.current === "edit") return true

					gridRef.current?.focus()

					return false
				}}
			>
				<AlertDialogHeader>
					<AlertDialogTitle>{alert?.title ?? t("previewSheetValidationAlertTitle")}</AlertDialogTitle>
					<AlertDialogDescription className="whitespace-pre-wrap">
						{alert?.message ?? t("previewSheetValidationAlertBody")}
						{alert?.style === "warning" ? `\n${t("previewSheetValidationContinue")}` : null}
					</AlertDialogDescription>
				</AlertDialogHeader>
				<AlertDialogFooter>
					{(alert === undefined ? [] : ALERT_BUTTONS[alert.style]).map((button, index) => (
						<Button
							key={button.key}
							type="button"
							variant={index === 0 ? "default" : "outline"}
							autoFocus={index === 0}
							onClick={() => {
								answer(button.verdict)
							}}
						>
							{button.key === "cancel" ? t("common:cancel") : t(button.key)}
						</Button>
					))}
				</AlertDialogFooter>
			</AlertDialogContent>
		</AlertDialog>
	)

	return { ask, dialog }
}
