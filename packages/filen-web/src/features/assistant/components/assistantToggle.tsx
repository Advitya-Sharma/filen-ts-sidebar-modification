import { useTranslation } from "react-i18next"
import { SparklesIcon } from "lucide-react"
import { cn } from "@filen/shared"
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip"
import { useAction } from "@/lib/keymap/useAction"
import { isAnyDialogOpen } from "@/lib/keymap/dialogGuard"
import { Kbd } from "@/lib/keymap/kbd"
import { useIsNarrowViewport } from "@/features/shell/hooks/useIsNarrowViewport"
import { RailProgressBar, RailTooltipDetail } from "@/features/shell/components/railProgressBar"
import {
	hydrateAssistantPanelOpen,
	toggleAssistant,
	useAssistantStore,
	useAssistantVisible
} from "@/features/assistant/store/useAssistantStore"
import { downloadingPercent } from "@/features/assistant/components/assistant.logic"
import { useResumeInterruptedDownload } from "@/features/assistant/components/assistantDownloadResume"
import { ASSISTANT_PANEL_ID, ASSISTANT_TOGGLE_ID, moveFocusOutOfAssistant } from "@/features/assistant/components/assistantFocus"

// The rail is mounted for the whole authed session, so the persisted open state loads once, here.
void hydrateAssistantPanelOpen()

// Opens the inline panel at desktop widths and the drawer below them; pressed while the panel shows. The
// rail owns the visual recipe (`itemClass`), so the toggle looks like every other rail entry. A model
// download shows on it as the Transfers entry shows its own, panel open or not: it can run unasked (resumed).
export function AssistantToggle({ itemClass }: { itemClass: (active: boolean) => string }) {
	const { t } = useTranslation(["common", "assistant"])
	const narrow = useIsNarrowViewport()
	const visible = useAssistantVisible(narrow)
	// A whole percentage, so progress re-renders the button once per percent at most.
	const percent = useAssistantStore(state => downloadingPercent(state.modelState))

	useResumeInterruptedDownload()

	function toggle(): void {
		if (visible && !narrow) {
			moveFocusOutOfAssistant()
		}

		toggleAssistant(narrow)
	}

	// Like the rail's route-nav actions, an open dialog owns the keyboard.
	useAction(
		"app.toggleAssistant",
		() => {
			if (isAnyDialogOpen()) {
				return
			}

			toggle()
		},
		undefined,
		[narrow, visible]
	)

	return (
		<Tooltip>
			<TooltipTrigger
				render={
					<button
						type="button"
						id={ASSISTANT_TOGGLE_ID}
						aria-label={percent === null ? t("toggleAssistant") : t("assistant:assistantToggleDownloading")}
						aria-pressed={visible}
						aria-controls={visible ? ASSISTANT_PANEL_ID : undefined}
						className={percent === null ? itemClass(visible) : cn(itemClass(visible), "relative")}
						onClick={toggle}
					>
						<SparklesIcon />
						{percent === null ? null : <RailProgressBar percent={percent} />}
					</button>
				}
			/>
			<TooltipContent side="right">
				{t("toggleAssistant")}
				{percent === null ? null : <RailTooltipDetail>{t("assistant:assistantStatusDownloading", { percent })}</RailTooltipDetail>}
				<Kbd action="app.toggleAssistant" />
			</TooltipContent>
		</Tooltip>
	)
}
