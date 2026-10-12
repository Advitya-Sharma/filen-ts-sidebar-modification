import type { DocxEditContext } from "./context"

// Document protection modes that allow editing within limits (w:documentProtection w:edit="forms",
// "comments", "trackedChanges"). Until they do, any enforced protection keeps the document view-only (gate 0
// reason "protected").

export type DocxProtectionMode = "none" | "forms" | "comments" | "trackedChanges"

export interface DocxProtection {
	readonly mode: DocxProtectionMode
	// Tracking cannot be switched off (trackedChanges).
	readonly forcesTracking: boolean
}

export const NO_PROTECTION: DocxProtection = { mode: "none", forcesTracking: false }

export function protectionOf(_context: DocxEditContext): DocxProtection {
	return NO_PROTECTION
}
