import { useSyncExternalStore } from "react"
import type { PasswordVerdict } from "@/features/preview/components/docxPasswordPrompt"

// The password a protected file's Download as asks, one file at a time: the export waits on ask() while the one
// prompt (downloadAsPasswordDialog.tsx, mounted at the root beside the toasts) checks passwords with `verify`. The
// typed password goes to `verify` only; nothing here keeps it.

// What the password goes on to protect: the PDF made from the file, or its copy converted to another format.
export type DownloadAsPasswordPurpose = "pdf" | "copy"

export interface DownloadAsPasswordAsk {
	readonly name: string
	readonly purpose: DownloadAsPasswordPurpose
	readonly verify: (password: string, signal: AbortSignal) => Promise<PasswordVerdict>
	// The prompt went: accepted (a password verified) or not (cancelled, or the export was).
	readonly settle: (accepted: boolean) => void
}

let current: DownloadAsPasswordAsk | null = null
const listeners = new Set<() => void>()

function publish(next: DownloadAsPasswordAsk | null): void {
	current = next

	for (const listener of listeners) {
		listener()
	}
}

function subscribe(listener: () => void): () => void {
	listeners.add(listener)

	return () => {
		listeners.delete(listener)
	}
}

export function currentPasswordAsk(): DownloadAsPasswordAsk | null {
	return current
}

export function usePasswordAsk(): DownloadAsPasswordAsk | null {
	return useSyncExternalStore(subscribe, currentPasswordAsk, currentPasswordAsk)
}

// What a check found, as the prompt takes it, or "refused": no password helps (the file cannot be opened, or the
// password cannot protect a PDF); the prompt closes and the export reports why.
export type DownloadAsPasswordVerdict = PasswordVerdict | "refused"

// The asks waiting their turn behind the one shown.
let queue: Promise<unknown> = Promise.resolve()

// Asks for `name`'s password until `verify` accepts one (true) or refuses the file, the prompt is cancelled, or
// `signal` (the export's Cancel) aborts (false). One prompt at a time: an ask made while another shows waits for it.
export function askPassword(
	name: string,
	verify: (password: string, signal: AbortSignal) => Promise<DownloadAsPasswordVerdict>,
	signal: AbortSignal,
	purpose: DownloadAsPasswordPurpose = "pdf"
): Promise<boolean> {
	const asked = queue.then(() => showAsk(name, verify, signal, purpose))

	queue = asked

	return asked
}

function showAsk(
	name: string,
	verify: (password: string, signal: AbortSignal) => Promise<DownloadAsPasswordVerdict>,
	signal: AbortSignal,
	purpose: DownloadAsPasswordPurpose
): Promise<boolean> {
	if (signal.aborted) {
		return Promise.resolve(false)
	}

	return new Promise(resolve => {
		let settled = false
		const ask: DownloadAsPasswordAsk = {
			name,
			purpose,
			verify: async (password, check) => {
				const verdict = await verify(password, check)

				if (verdict === "accepted" || verdict === "refused") {
					ask.settle(verdict === "accepted")
				}

				return verdict === "refused" ? "aborted" : verdict
			},
			settle: accepted => {
				if (settled) {
					return
				}

				settled = true
				signal.removeEventListener("abort", onAbort)

				if (current === ask) {
					publish(null)
				}

				resolve(accepted)
			}
		}
		const onAbort = (): void => {
			ask.settle(false)
		}

		signal.addEventListener("abort", onAbort, { once: true })
		publish(ask)
	})
}
