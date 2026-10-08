import { useEffect } from "react"
import { toast } from "sonner"
import { i18n } from "@/lib/i18n"
import { useAssistantStore } from "@/features/assistant/store/useAssistantStore"
import { cancelAssistantDownload, resumeAssistantDownload } from "@/features/assistant/store/assistantSession"
import { downloadByteFigures, resumedDownloadEnded, resumedDownloadStarted } from "@/features/assistant/components/assistant.logic"

const RESUMED_TOAST_ID = "assistant-download-resumed"

// For the lifetime of the signed-in shell: asks the runtime to continue a download a reload, crash or closed
// tab cut off, and says so when it does, with the way to stop it, since no click started it this time.
export function useResumeInterruptedDownload(): void {
	useEffect(() => {
		const unsubscribe = useAssistantStore.subscribe((state, previous) => {
			const next = state.modelState

			// Its Cancel stops whatever download runs: the toast goes with the download it announced, even
			// while a hidden or hovered page holds its timer.
			if (resumedDownloadEnded(previous.modelState, next)) {
				toast.dismiss(RESUMED_TOAST_ID)

				return
			}

			if (!resumedDownloadStarted(previous.modelState, next)) {
				return
			}

			toast(i18n.t("assistant:assistantDownloadResumedToast"), {
				id: RESUMED_TOAST_ID,
				description: i18n.t(
					"assistant:assistantDownloadResumedToastBody",
					downloadByteFigures(next.receivedBytes, next.model.sizeBytes)
				),
				action: { label: i18n.t("assistant:assistantCancelDownload"), onClick: cancelAssistantDownload }
			})
		})

		resumeAssistantDownload()

		return unsubscribe
	}, [])
}
