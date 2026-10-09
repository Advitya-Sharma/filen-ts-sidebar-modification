// The newest generation the service worker has seen for each port-registration id, admitted or released,
// so a registration that lands behind a newer one, or behind a release that already covered it, is refused
// (SW_MSG_REGISTER_PORT_PREVIEW). Bounded: only a recent id can still have a registration on its way.
export class GenerationGate {
	private readonly newest = new Map<string, number>()
	private readonly max: number

	public constructor(max: number) {
		this.max = max
	}

	// True, and recorded, when `generation` is newer than anything seen for `id`.
	public admit(id: string, generation: number): boolean {
		const seen = this.newest.get(id)

		if (seen !== undefined && generation <= seen) {
			return false
		}

		this.record(id, generation)

		return true
	}

	// A release of every generation of `id` up to `generation`.
	public retire(id: string, generation: number): void {
		this.record(id, Math.max(generation, this.newest.get(id) ?? generation))
	}

	private record(id: string, generation: number): void {
		// Re-inserted so the bound drops the least recently touched id.
		this.newest.delete(id)
		this.newest.set(id, generation)

		if (this.newest.size > this.max) {
			const oldest = this.newest.keys().next().value

			if (oldest !== undefined) {
				this.newest.delete(oldest)
			}
		}
	}
}
