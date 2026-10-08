import { type } from "arktype"

// The user's opt-in to a download, kept past the page while that download runs: removed once it completes,
// fails, is cancelled or the model is removed. Still set at the next start, it means a reload, a crash or a
// closed tab cut the download off, which may then continue without a click, for the account that asked
// only (the model files are device-wide, the opt-in is not). Sign-out clears it with the rest of the work
// in flight. Storage blocked: nothing resumes.
const DOWNLOAD_INTENT_KEY = "filen.assistant.downloadIntent.v1"

// Automatic starts in a row that found no more on disk than the one before: the download itself may be
// what takes the page down, so it waits for a click instead of crashing every start.
export const MAX_FRUITLESS_AUTO_STARTS = 2

const downloadIntentSchema = type({
	modelId: "string",
	account: "string",
	// Automatic starts since the last click or the last one that found progress.
	autoStarts: "number.integer >= 0",
	// On disk when the last start began.
	presentBytes: "number >= 0"
})

export type DownloadIntent = typeof downloadIntentSchema.infer

export function readDownloadIntent(): DownloadIntent | null {
	try {
		const raw = localStorage.getItem(DOWNLOAD_INTENT_KEY)

		if (raw === null) {
			return null
		}

		const parsed = downloadIntentSchema(JSON.parse(raw))

		return parsed instanceof type.errors ? null : parsed
	} catch {
		return null
	}
}

export function writeDownloadIntent(intent: DownloadIntent | null): void {
	try {
		if (intent === null) {
			localStorage.removeItem(DOWNLOAD_INTENT_KEY)
		} else {
			localStorage.setItem(DOWNLOAD_INTENT_KEY, JSON.stringify(intent))
		}
	} catch {
		// Storage blocked: an interrupted download waits for a click instead.
	}
}

// The record for an automatic start, or null when this one must wait for a click: another model's or
// another account's opt-in, or too many starts in a row that found no progress.
export function autoStartIntent(
	intent: DownloadIntent | null,
	{ modelId, account, presentBytes }: { modelId: string; account: string | null; presentBytes: number }
): DownloadIntent | null {
	if (intent?.modelId !== modelId || account === null || intent.account !== account) {
		return null
	}

	const autoStarts = presentBytes > intent.presentBytes ? 0 : intent.autoStarts

	return autoStarts >= MAX_FRUITLESS_AUTO_STARTS ? null : { modelId, account, autoStarts: autoStarts + 1, presentBytes }
}
