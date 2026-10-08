import { useTranslation } from "react-i18next"
import type { TranscriptStep } from "@/features/assistant/store/transcript"
import {
	ROUTE_KEYS,
	STEP_FAILURE_KEYS,
	durationFigure,
	stepCountNoun,
	toolActivityKey,
	toolSubjectKey
} from "@/features/assistant/components/assistant.logic"

export interface StepText {
	label: string
	// What came of it, or why it stopped; null when the glyph says it all.
	detail: string | null
}

const secondFormats = new Map<string, Intl.NumberFormat>()

function secondFormat(locale: string, fractionDigits: 0 | 1): Intl.NumberFormat {
	const key = `${locale}|${String(fractionDigits)}`
	let format = secondFormats.get(key)

	if (format === undefined) {
		format = new Intl.NumberFormat(locale, { minimumFractionDigits: fractionDigits, maximumFractionDigits: fractionDigits })
		secondFormats.set(key, format)
	}

	return format
}

// "4.2 s"; a running step counts whole seconds.
export function useDurationText(): (ms: number, live: boolean) => string {
	const { t, i18n } = useTranslation("assistant")

	return (ms, live) => {
		const { seconds, fractionDigits } = durationFigure(ms, live)

		return t("assistantDurationSeconds", { seconds: secondFormat(i18n.language, fractionDigits).format(seconds) })
	}
}

// A step's wording, also what the live region reads. A subject is plain text from the runtime (sanitized,
// shortened) and only ever interpolated into a text node.
export function useStepText(): (step: TranscriptStep) => StepText {
	const { t } = useTranslation("assistant")

	const count = (tool: string, resultCount: number, more: boolean): string => {
		switch (stepCountNoun(tool)) {
			case "notes":
				return t("assistantStepCountNotes", { count: resultCount })
			case "chats":
				return t("assistantStepCountChats", { count: resultCount })
			case "playlists":
				return t("assistantStepCountPlaylists", { count: resultCount })
			case "items":
				return more
					? t("assistantStepCountItemsMore", { count: resultCount })
					: t("assistantStepCountItems", { count: resultCount })
		}
	}

	const ended = (status: TranscriptStep["status"]): string | null => (status === "stopped" ? t("assistantStepStopped") : null)

	return step => {
		switch (step.kind) {
			case "loading":
				return {
					label: t(
						step.status === "done"
							? "assistantStepLoaded"
							: step.status === "failed"
								? "assistantStepLoadFailed"
								: "assistantStepLoading"
					),
					detail: ended(step.status)
				}
			case "routing":
				return {
					label: step.status === "done" && step.route !== null ? t(ROUTE_KEYS[step.route]) : t("assistantStepRouting"),
					detail: ended(step.status)
				}
			case "thinking":
				return { label: t("assistantStepThinking"), detail: ended(step.status) }
			case "replying":
				return { label: t("assistantStepReplying"), detail: ended(step.status) }
			case "tool": {
				const subjectKey = toolSubjectKey(step.tool)
				const label =
					subjectKey !== null && step.subject !== null ? t(subjectKey, { subject: step.subject }) : t(toolActivityKey(step.tool))
				let detail: string | null

				if (step.status === "waiting") {
					detail = t("assistantStepWaiting")
				} else if (step.status === "queued") {
					detail = t("assistantStepQueued")
				} else if (step.status === "failed") {
					detail = t(STEP_FAILURE_KEYS[step.failure ?? "error"])
				} else if (step.status === "done" && step.resultCount !== null) {
					detail = count(step.tool, step.resultCount, step.resultMore)
				} else {
					detail = ended(step.status)
				}

				return { label, detail }
			}
		}
	}
}
