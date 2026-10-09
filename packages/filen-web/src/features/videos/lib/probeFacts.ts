import { kvEntriesJson } from "@/lib/storage/adapter"
import { PROBE_KV_PREFIX, probeRecordSchema, type ProbeRecord } from "@/features/videos/lib/probeService.logic"
import { probeService } from "@/features/videos/lib/probeService"
import type { ProbeFacts } from "@/features/videos/lib/videoGrid.logic"

// Runtime and picture height per file uuid, for the grids' runtime sort and resolution filter: every stored probe
// record read once per session (only when a grid first needs them), then kept current from the probes this
// session runs. Never probes.

function factsOf(record: ProbeRecord): ProbeFacts {
	return {
		durationSec: record.probe?.durationSec ?? null,
		height: record.probe?.video?.h ?? 0
	}
}

class ProbeFactsStore {
	private facts: Map<string, ProbeFacts> | null = null
	private loading: Promise<ReadonlyMap<string, ProbeFacts>> | null = null
	private generation = 0

	public constructor() {
		probeService.onRecord((uuid, record) => {
			if (this.facts !== null) {
				this.facts = new Map(this.facts).set(uuid, factsOf(record))
			}
		})
	}

	public load(): Promise<ReadonlyMap<string, ProbeFacts>> {
		if (this.facts !== null) {
			return Promise.resolve(this.facts)
		}

		const generation = this.generation

		this.loading ??= kvEntriesJson(PROBE_KV_PREFIX, probeRecordSchema)
			.then(rows => {
				const facts = new Map<string, ProbeFacts>()

				for (const [key, record] of rows) {
					facts.set(key.startsWith(PROBE_KV_PREFIX) ? key.slice(PROBE_KV_PREFIX.length) : key, factsOf(record))
				}

				if (generation === this.generation) {
					this.facts = facts
				}

				return facts
			})
			.finally(() => {
				this.loading = null
			})

		return this.loading
	}

	// Logout.
	public reset(): void {
		this.generation++
		this.facts = null
		this.loading = null
	}
}

export const probeFacts = new ProbeFactsStore()
