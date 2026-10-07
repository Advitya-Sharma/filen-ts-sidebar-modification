import { useTranslation } from "react-i18next"
import { SparklesIcon } from "lucide-react"
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip"
import { useAction } from "@/lib/keymap/useAction"
import { isAnyDialogOpen } from "@/lib/keymap/dialogGuard"
import { Kbd } from "@/lib/keymap/kbd"
import { useIsNarrowViewport } from "@/features/shell/hooks/useIsNarrowViewport"
import { hydrateAssistantPanelOpen, toggleAssistant, useAssistantVisible } from "@/features/assistant/store/useAssistantStore"
import { ASSISTANT_PANEL_ID, ASSISTANT_TOGGLE_ID, moveFocusOutOfAssistant } from "@/features/assistant/components/assistantFocus"

// The rail is mounted for the whole authed session, so the persisted open state loads once, here.
void hydrateAssistantPanelOpen()

// Opens the inline panel at desktop widths and the drawer below them; pressed while the panel shows. The
// rail owns the visual recipe (`itemClass`), so the toggle looks like every other rail entry.
export function AssistantToggle({ itemClass }: { itemClass: (active: boolean) => string }) {
	const { t } = useTranslation()
	const narrow = useIsNarrowViewport()
	const visible = useAssistantVisible(narrow)

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
						aria-label={t("toggleAssistant")}
						aria-pressed={visible}
						aria-controls={visible ? ASSISTANT_PANEL_ID : undefined}
						className={itemClass(visible)}
						onClick={toggle}
					>
						<SparklesIcon />
					</button>
				}
			/>
			<TooltipContent side="right">
				{t("toggleAssistant")}
				<Kbd action="app.toggleAssistant" />
			</TooltipContent>
		</Tooltip>
	)
}
