import { useTranslation } from "react-i18next"
import { SparklesIcon } from "lucide-react"
import type { AssistantKey } from "@/lib/i18n"
import { useAssistantStore } from "@/features/assistant/store/useAssistantStore"
import { sendAssistantMessage } from "@/features/assistant/store/assistantSession"
import type { TranscriptEntry } from "@/features/assistant/store/transcript"
import { AssistantTranscript } from "@/features/assistant/components/assistantTranscript"
import { TURN_END_KEYS, spokenText, transcriptRows } from "@/features/assistant/components/assistant.logic"
import { useProposalTitle } from "@/features/assistant/components/useProposalTitle"
import { AssistantComposer } from "@/features/assistant/components/assistantComposer"
import { AssistantModelStatus } from "@/features/assistant/components/assistantModelStatus"

const EXAMPLE_KEYS: readonly AssistantKey[] = [
	"assistantExampleFind",
	"assistantExampleSummarize",
	"assistantExampleOpen",
	"assistantExamplePlay"
]

function EmptyConversation() {
	const { t } = useTranslation("assistant")

	return (
		<div className="flex min-h-0 flex-1 flex-col items-center justify-center gap-3 overflow-y-auto px-6 py-8 text-center">
			<div
				aria-hidden="true"
				className="flex size-12 items-center justify-center rounded-2xl bg-muted text-muted-foreground"
			>
				<SparklesIcon className="size-6" />
			</div>
			<h3 className="text-base font-semibold">{t("assistantEmptyTitle")}</h3>
			<p className="text-sm text-muted-foreground">{t("assistantEmptyBody")}</p>
			<ul
				aria-label={t("assistantExamplesLabel")}
				className="flex w-full max-w-80 flex-col gap-1.5"
			>
				{EXAMPLE_KEYS.map(key => (
					<li key={key}>
						<button
							type="button"
							className="w-full rounded-xl border border-border bg-background/60 px-3 py-2 text-left text-sm focus-ring transition-colors outline-none hover:bg-muted"
							onClick={() => {
								sendAssistantMessage(t(key))
							}}
						>
							{t(key)}
						</button>
					</li>
				))}
			</ul>
		</div>
	)
}

// One polite announcement when a reply starts and one when its turn settles, never per token: the
// streaming bubble itself is not a live region.
function Announcer({ entries }: { entries: readonly TranscriptEntry[] }) {
	const { t } = useTranslation("assistant")
	const generating = useAssistantStore(state => state.activeTurnId !== null)
	const settled = useAssistantStore(state => state.settled)
	const proposalTitle = useProposalTitle()
	let text = ""

	if (generating) {
		text = t("assistantResponding")
	} else if (settled !== null) {
		const parts: string[] = []

		for (const entry of entries) {
			if (entry.turnId !== settled.turnId) {
				continue
			}

			if (entry.kind === "reply") {
				parts.push(spokenText(entry.text))
			} else if (entry.kind === "proposal" && entry.status === "pending") {
				parts.push(t("assistantProposalAnnouncement", { title: proposalTitle(entry.proposal) }))
			} else if (entry.kind === "turnEnd") {
				parts.push(t(TURN_END_KEYS[entry.outcome]))
			}
		}

		text = parts.join("\n")
	}

	return (
		<div
			aria-live="polite"
			className="sr-only"
		>
			{text}
		</div>
	)
}

export function AssistantConversation() {
	const entries = useAssistantStore(state => state.entries)
	const generating = useAssistantStore(state => state.activeTurnId !== null)

	return (
		<>
			{entries.length === 0 ? <EmptyConversation /> : <AssistantTranscript rows={transcriptRows(entries, generating)} />}
			<Announcer entries={entries} />
			<AssistantModelStatus />
			<AssistantComposer />
		</>
	)
}
