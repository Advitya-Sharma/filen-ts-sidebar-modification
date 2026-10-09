import { onlineManager } from "@tanstack/react-query"
import type { AnyFile } from "@filen/sdk-rs"
import { sdkApi } from "@/lib/sdk/client"
import { log } from "@/lib/log"
import { RankedGate } from "@/features/drive/lib/rankedGate"
import { rangeReads, type RangeReadBudget } from "@/lib/media/rangeReads"
import type { ProbeTier, VideoProbe } from "@/lib/media/probe/types"
import type { RecordStore } from "@/features/videos/lib/recordStore.logic"
import { probeRecords } from "@/features/videos/lib/records"
import {
	PROBE_CONCURRENCY,
	PROBE_DROP_GRACE_MS,
	PROBE_FAILURE_LIMIT,
	PROBE_FINISH_WITHIN_BLOCKS,
	mayStartProbe,
	mergeProbe,
	recordSatisfies,
	type ProbeRecord
} from "@/features/videos/lib/probeService.logic"

// Probes video files on demand, cheapest tier first, each tier at most once per content version: every
// record is persisted per file uuid, and a request a record already answers costs no download. The SDK
// owns network concurrency and retries; the gates here bound how many probes this page keeps in flight
// and which goes next.

// Lets the visible cards mounting in the same pass claim free slots before the rest of the queue.
const PROBE_HOLD_MS = 50

export interface ProbeTarget {
	uuid: string
	file: AnyFile
}

export interface ProbeRequestOptions {
	// How soon this caller wants it (PROBE_RANK), read whenever a slot frees; explicit when absent.
	rank?: () => number
	// Scopes this caller's interest: once every interested caller has aborted, the probe is dropped
	// (after a grace). Without one the probe always runs.
	signal?: AbortSignal
}

export interface ProbeServiceDeps {
	readProbe: (file: AnyFile, uuid: string, tier: ProbeTier, token: string) => Promise<VideoProbe | null>
	// Aborts the probe's reads (or the probe about to start under `token`), unless `finishWithin` is given
	// and the probe is that many blocks from the end of its current read budget.
	cancel: (token: string, finishWithin?: number) => void
	records: RecordStore<ProbeRecord>
	reads: Pick<RangeReadBudget, "acquire" | "release">
	isOnline: () => boolean
	now: () => number
}

interface Claim {
	interest: number
	ranks: Set<() => number>
	pinned: boolean
	controller: AbortController
	dropTimer: ReturnType<typeof setTimeout> | undefined
	settled: boolean
}

interface Pending {
	promise: Promise<ProbeRecord | null>
	claim: Claim
	tier: ProbeTier
}

function explicitRank(): number {
	return 0
}

export class ProbeService {
	private readonly deps: ProbeServiceDeps
	private readonly gate = new RankedGate(PROBE_CONCURRENCY, PROBE_HOLD_MS)
	// uuid → the probe in flight for it (one per file: a deeper tier joins nothing shallower).
	private readonly pending = new Map<string, Pending>()
	private readonly failures = new Map<string, number>()
	private readonly tokens = new Set<string>()
	private started = 0
	private epoch = 0

	public constructor(deps: ProbeServiceDeps) {
		this.deps = deps
	}

	// Probes started this session.
	public get startedCount(): number {
		return this.started
	}

	// The record in memory, without I/O: undefined when not looked up yet.
	public peek(uuid: string): ProbeRecord | null | undefined {
		return this.deps.records.peek(uuid)
	}

	// The record answering `tier` for the file: from memory or kv, else from a probe. Null when none can be
	// had now: offline, past the file's failure limit, past the session cap for a non-explicit request,
	// dropped because nobody wanted it, or a failed download. Never throws.
	public async request(target: ProbeTarget, tier: ProbeTier, options: ProbeRequestOptions = {}): Promise<ProbeRecord | null> {
		const known = await this.deps.records.get(target.uuid)

		if (known !== null && recordSatisfies(known, tier)) {
			return known
		}

		if (options.signal?.aborted === true) {
			return null
		}

		const existing = this.pending.get(target.uuid)

		if (existing !== undefined && !existing.claim.controller.signal.aborted && (existing.tier === tier || existing.tier === "full")) {
			this.registerInterest(existing.claim, options)

			return existing.promise
		}

		if ((this.failures.get(target.uuid) ?? 0) >= PROBE_FAILURE_LIMIT || !this.deps.isOnline()) {
			return null
		}

		const claim: Claim = {
			interest: 0,
			ranks: new Set(),
			pinned: false,
			controller: new AbortController(),
			dropTimer: undefined,
			settled: false
		}
		// A shallower probe still in flight finishes first, so this one starts from its record; it queues at
		// this request's rank meanwhile, so an explicit request never waits behind a card's.
		const before = existing?.promise ?? Promise.resolve(null)

		existing?.claim.ranks.add(() => this.claimRank(claim))
		const entry: Pending = { promise: before.then(() => this.run(target, tier, claim, this.epoch)), claim, tier }
		const settle = (): void => {
			claim.settled = true
			clearTimeout(claim.dropTimer)

			if (this.pending.get(target.uuid) === entry) {
				this.pending.delete(target.uuid)
			}
		}

		this.pending.set(target.uuid, entry)
		this.registerInterest(claim, options)
		void entry.promise.then(settle, settle)

		return entry.promise
	}

	private registerInterest(claim: Claim, { signal, rank }: ProbeRequestOptions): void {
		if (signal === undefined) {
			claim.pinned = true

			if (rank !== undefined) {
				claim.ranks.add(rank)
			}

			return
		}

		claim.interest++
		clearTimeout(claim.dropTimer)

		const ranked = rank ?? explicitRank

		claim.ranks.add(ranked)
		signal.addEventListener(
			"abort",
			() => {
				claim.interest--
				claim.ranks.delete(ranked)

				if (claim.settled || claim.pinned || claim.interest > 0) {
					return
				}

				clearTimeout(claim.dropTimer)
				claim.dropTimer = setTimeout(() => {
					if (!claim.pinned && claim.interest <= 0) {
						claim.controller.abort()
					}
				}, PROBE_DROP_GRACE_MS)
			},
			{ once: true }
		)
	}

	private unwanted(claim: Claim): boolean {
		return claim.controller.signal.aborted || (!claim.pinned && claim.interest <= 0)
	}

	// Work nobody wants goes first, to hand its slot straight back; otherwise the nearest caller's rank.
	private claimRank(claim: Claim): number {
		if (this.unwanted(claim)) {
			return Number.NEGATIVE_INFINITY
		}

		if (claim.ranks.size === 0) {
			return 0
		}

		let best = Number.POSITIVE_INFINITY

		for (const rank of claim.ranks) {
			best = Math.min(best, rank())
		}

		return best
	}

	private async run(target: ProbeTarget, tier: ProbeTier, claim: Claim, epoch: number): Promise<ProbeRecord | null> {
		const rank = (): number => this.claimRank(claim)

		await this.gate.acquire(rank)

		try {
			if (this.unwanted(claim) || epoch !== this.epoch || !mayStartProbe(rank(), this.started)) {
				return null
			}

			// Another tab may have probed the file while this one queued.
			const known = await this.deps.records.refresh(target.uuid)

			if (known !== null && recordSatisfies(known, tier)) {
				return known
			}

			await this.deps.reads.acquire(rank)

			try {
				if (this.unwanted(claim) || epoch !== this.epoch) {
					return null
				}

				return await this.probe(target, tier, claim, epoch)
			} finally {
				this.deps.reads.release()
			}
		} finally {
			this.gate.release()
		}
	}

	private async probe(target: ProbeTarget, tier: ProbeTier, claim: Claim, epoch: number): Promise<ProbeRecord | null> {
		const token = crypto.randomUUID()
		const stop = (): void => {
			this.deps.cancel(token, PROBE_FINISH_WITHIN_BLOCKS)
		}

		this.started++
		this.tokens.add(token)
		claim.controller.signal.addEventListener("abort", stop, { once: true })

		try {
			const probe = await this.deps.readProbe(target.file, target.uuid, tier, token)

			if (epoch !== this.epoch) {
				return null
			}

			const record = mergeProbe(await this.deps.records.refresh(target.uuid), tier, probe, this.deps.now())

			this.deps.records.put(target.uuid, record)

			return record
		} catch (error) {
			if (epoch === this.epoch && !claim.controller.signal.aborted) {
				log.warn("videos", "probe failed", target.uuid, tier, error)
				this.failures.set(target.uuid, (this.failures.get(target.uuid) ?? 0) + 1)
			}

			return null
		} finally {
			claim.controller.signal.removeEventListener("abort", stop)
			this.tokens.delete(token)
		}
	}

	// Coming back online gives every file that failed a fresh budget.
	public onReconnect(): void {
		this.failures.clear()
	}

	// Logout: abort what runs, forget every verdict and record held in memory.
	public reset(): void {
		this.epoch++

		for (const entry of this.pending.values()) {
			entry.claim.controller.abort()
		}

		for (const token of this.tokens) {
			this.deps.cancel(token)
		}

		this.pending.clear()
		this.tokens.clear()
		this.failures.clear()
		this.started = 0
		this.deps.records.reset()
	}
}

export const probeService = new ProbeService({
	readProbe: (file, uuid, tier, token) => sdkApi.readVideoProbe(file, uuid, tier, token),
	cancel: (token, finishWithin) => {
		void (finishWithin === undefined
			? sdkApi.cancelPreviewDownload(token)
			: sdkApi.cancelPreviewDownloadUnlessFinishing(token, finishWithin))
	},
	records: probeRecords,
	reads: rangeReads,
	isOnline: () => onlineManager.isOnline(),
	now: () => Date.now()
})

// Module scope and without React, like the thumbnail service's reconnect hook: it lives as long as the tab.
onlineManager.subscribe(online => {
	if (online) {
		probeService.onReconnect()
	}
})
