import { registerAssistantSession } from "@/features/assistant/store/assistantSession"
import { AssistantRuntime } from "@/features/assistant/runtime/client"
import { createAssistantSession } from "@/features/assistant/loop/session"
import { createAssistantToolExecutor } from "@/features/assistant/tools/executor"

// Importing this module (main.tsx) plugs the on-device runtime into the assistant UI. Nothing heavy
// happens here: once the page is idle the runtime checks the browser's features and lists the model's
// OPFS file sizes. The worker starts only on a download or when the panel opens with a downloaded model.
export const assistantRuntime = new AssistantRuntime()

export const assistantSession = createAssistantSession({ runtime: assistantRuntime, executor: createAssistantToolExecutor() })

registerAssistantSession(assistantSession)

const check = () => {
	void assistantRuntime.init()
}

if (typeof requestIdleCallback === "function") {
	requestIdleCallback(check, { timeout: 5000 })
} else {
	setTimeout(check, 2000)
}
