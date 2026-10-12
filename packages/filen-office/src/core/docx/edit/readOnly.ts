import type { Story } from "../model/index"
import type { DocxGateReason } from "../package"
import type { DocxEditContext } from "./context"

// "textbox": text box stories are view-only until editing them also updates the VML fallback Word keeps
// of each text box (mc:Fallback), which would otherwise show the old text.
export type DocxReadOnlyReason = DocxGateReason | "host" | "textbox"

// Why a session of `story` is read-only, in order: the document's gate reasons, the host's mode, the story's
// own. Empty when it may be edited.
export function readOnlyReasons(context: DocxEditContext, story: Story, host: boolean): readonly DocxReadOnlyReason[] {
	const out: DocxReadOnlyReason[] = [...context.gate]

	if (host) out.push("host")
	if (story.kind === "textbox") out.push("textbox")

	return out
}
