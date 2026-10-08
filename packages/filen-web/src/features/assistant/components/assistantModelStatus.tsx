import { useTranslation } from "react-i18next"
import { Button } from "@/components/ui/button"
import { Spinner } from "@/components/ui/spinner"
import { useAssistantStore } from "@/features/assistant/store/useAssistantStore"
import { loadAssistantModel } from "@/features/assistant/store/assistantSession"
import { assistantInitialFocusTarget } from "@/features/assistant/components/assistantFocus"
import { modelShowsLoading } from "@/features/assistant/components/assistant.logic"

// The Load button leaves with the state it belongs to, which would drop focus to the page: the composer,
// where the user goes next, takes it first.
function loadFromButton(): void {
	assistantInitialFocusTarget()?.focus()
	loadAssistantModel()
}

// One quiet line above the composer while the model is not in memory: sending loads it, so this only says
// so (and why, after an idle unload) and offers loading it ahead of time.
export function AssistantModelStatus() {
	const { t } = useTranslation("assistant")
	const modelState = useAssistantStore(state => state.modelState)
	const generating = useAssistantStore(state => state.activeTurnId !== null)

	if (modelShowsLoading(modelState, generating)) {
		return (
			<p className="flex shrink-0 items-center gap-2 px-5 pt-1 text-xs leading-5 text-muted-foreground">
				<Spinner
					aria-hidden="true"
					className="size-3.5"
				/>
				<span role="status">
					{modelState.kind === "loading" && modelState.verifying ? t("assistantVerifyingTitle") : t("assistantLoadingTitle")}
				</span>
			</p>
		)
	}

	if (modelState.kind !== "downloaded") {
		return null
	}

	return (
		<div className="flex shrink-0 items-center gap-2 pt-1 pr-3 pl-5 text-xs leading-5 text-muted-foreground">
			<p
				role="status"
				className="min-w-0 flex-1"
			>
				{modelState.idleUnloaded ? t("assistantModelIdleUnloaded") : t("assistantModelNotLoaded")}
			</p>
			<Button
				variant="ghost"
				size="xs"
				className="shrink-0"
				onClick={loadFromButton}
			>
				{t("assistantLoadModel")}
			</Button>
		</div>
	)
}
