import { usePreventRemove } from "expo-router"
import { useTranslation } from "react-i18next"
import { run } from "@filen/shared"
import useDrivePreviewStore from "@/stores/useDrivePreview.store"
import events from "@/lib/events"
import prompts from "@/lib/prompts"
import alerts from "@/lib/alerts"
import logger from "@/lib/logger"

// Mounted by the drivePreview route. Every way out of the preview — the header close button, the
// Android hardware back button, and the iOS/Android swipe-back gesture — funnels through the route
// pop, so a single removal prevention intercepts them all. While the open editable text file has
// unsaved edits it prompts Save / Discard / Cancel before leaving; otherwise navigation proceeds
// untouched. (Only the text/code preview is editable + manual-save; notes auto-save, images/videos
// have no editor.) Renders nothing.
const UnsavedChangesGuard = () => {
	const { t } = useTranslation()
	const hasUnsavedEdits = useDrivePreviewStore(state => state.hasUnsavedEdits)

	// The router blocks the removal while this is true; repeat() re-dispatches the blocked action once
	// the user decides.
	usePreventRemove(hasUnsavedEdits, ({ repeat }) => {
		void (async () => {
			const promptResult = await run(async () => {
				return await prompts.confirm3({
					title: t("unsaved_changes"),
					message: t("unsaved_changes_message"),
					primaryText: t("save"),
					destructiveText: t("discard"),
					cancelText: t("cancel")
				})
			})

			if (!promptResult.success) {
				logger.error("drivePreview", "unsaved-changes prompt failed", {
					error: promptResult.error
				})
				alerts.error(promptResult.error)

				events.emit("drivePreviewDismissBlocked")

				return
			}

			// Cancel (or dismissing the alert) keeps the user on the preview.
			if (promptResult.data === "cancel") {
				events.emit("drivePreviewDismissBlocked")

				return
			}

			if (promptResult.data === "primary") {
				const saveEdits = useDrivePreviewStore.getState().saveEdits
				// saveEdits() reports an ordinary failure by returning false, but it is not a total
				// function: previewText's post-upload bookkeeping (cache invalidation and the
				// driveItemUpdated fan-out) runs outside its own error boundary, and EventEmitter3 does
				// not catch subscriber throws. A rejection escaping here would skip the
				// drivePreviewDismissBlocked emit below and strand the user in a preview whose close
				// button and dismiss gesture are both latched off — with no way out on iOS.
				const saveResult = await run(async () => (saveEdits ? await saveEdits() : false))

				if (!saveResult.success) {
					logger.error("drivePreview", "saving before dismissal threw", {
						error: saveResult.error
					})
					alerts.error(saveResult.error)

					events.emit("drivePreviewDismissBlocked")

					return
				}

				// Could not save (e.g. offline) — keep the user put; save() already surfaced any error.
				if (!saveResult.data) {
					events.emit("drivePreviewDismissBlocked")

					return
				}
			}

			// Saved, or the user chose to discard: clear the flag (which lifts the prevention) and let the
			// original navigation run.
			useDrivePreviewStore.getState().setHasUnsavedEdits(false)
			repeat()
		})()
	})

	return null
}

export default UnsavedChangesGuard
