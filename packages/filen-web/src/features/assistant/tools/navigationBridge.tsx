import { useAssistantNavigationBridge } from "@/features/assistant/tools/navigation"

// The same as calling useAssistantNavigationBridge, for a host that would rather mount a component.
export function AssistantNavigationBridge(): null {
	useAssistantNavigationBridge()

	return null
}
