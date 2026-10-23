import { sliceClock, type SliceClock, type Slicing } from "../slicing"

// A long job's turns (a worker's read, proof or save): given in slices, the thread's own between them (a macrotask,
// so a close queued for it is read); the job ends with RtfAborted once `signal` is aborted.
export interface RtfTurns {
	readonly signal: { readonly aborted: boolean }
	readonly slicing: Slicing
}

export class RtfAborted extends Error {
	constructor() {
		super("aborted")
		this.name = "RtfAborted"
	}
}

// One job's turns: one slice clock across its phases (a job is the turns of the calls it makes). `aborting`: the
// slicing for sliced work of others (parsing a part), whose pauses end the job once aborted.
export interface RtfJob extends RtfTurns {
	readonly clock: SliceClock
	readonly due: () => boolean
	readonly aborting: Slicing
}

export function jobOf(turns: RtfTurns | undefined): RtfJob | undefined {
	if (turns === undefined) return undefined

	const clock = sliceClock(turns.slicing)
	const job: RtfJob = {
		signal: turns.signal,
		slicing: clock,
		clock,
		due: () => clock.due,
		aborting: { pause: () => turn(job), sliceMs: clock.sliceMs, now: clock.now }
	}

	return job
}

export async function turn(job: RtfJob | undefined): Promise<void> {
	if (job === undefined) return

	await job.clock.pause()

	if (job.signal.aborted) throw new RtfAborted()
}
