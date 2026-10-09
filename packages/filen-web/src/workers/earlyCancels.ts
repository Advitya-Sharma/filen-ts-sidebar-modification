// Cancels that arrive before their read has registered: a page-side caller can give up while its call
// still waits in a queue there, and a cancel that finds nothing to abort would otherwise be lost while
// the read starts anyway. The token is remembered for a while, so its read starts aborted. Callers mint a
// token per read and never reuse one, so a remembered token never catches a later, wanted read.
const EARLY_CANCEL_TTL_MS = 60_000
const MAX_EARLY_CANCELS = 1024

export class EarlyCancels {
	// token → when it was cancelled, oldest first.
	private readonly tokens = new Map<string, number>()
	private readonly now: () => number

	public constructor(now: () => number = Date.now) {
		this.now = now
	}

	public add(token: string): void {
		this.tokens.delete(token)
		this.tokens.set(token, this.now())

		while (this.tokens.size > MAX_EARLY_CANCELS) {
			const oldest = this.tokens.keys().next()

			if (oldest.done === true) {
				break
			}

			this.tokens.delete(oldest.value)
		}
	}

	// Whether `token` was cancelled before its read registered; forgotten once asked.
	public take(token: string): boolean {
		const at = this.tokens.get(token)

		this.tokens.delete(token)

		return at !== undefined && this.now() - at <= EARLY_CANCEL_TTL_MS
	}
}
