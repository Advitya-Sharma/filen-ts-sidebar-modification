import { sdkApi } from "@/lib/sdk/client"
import { asErrorDTO } from "@/lib/sdk/errors"
import { assertSwControlled, triggerSwStreamDownload, type SwSaveTarget } from "@/features/drive/lib/saveDownload"

// The page's SDK call `produce` writes into a stream whose readable the service worker pipes into the browser's
// download (Content-Length `size`, none when null). The navigation waits for the first byte (or the end of an empty
// stream), so a call that fails before it leaves no failed download in the browser. A failure of either side stops
// the other. A cancel names the outcome whichever way it arrived (the browser's own cancel reaches the SDK as a
// failed write), unless it is the stop this sent; else the first failure does. `transferId` is the SDK call's.
export async function streamProducerToSw<T>(
	save: SwSaveTarget,
	transferId: string,
	size: number | null,
	produce: (writable: WritableStream<Uint8Array>) => Promise<T>
): Promise<T> {
	assertSwControlled()

	let flowing: () => void = () => undefined
	const firstChunk = new Promise<void>(resolve => {
		flowing = resolve
	})
	// Room for one chunk on the readable side: with none, nothing reaches transform() before the worker
	// reads, and the worker reads only after the navigation this waits on.
	const pipe = new TransformStream<Uint8Array, Uint8Array>(
		{
			transform(chunk, controller) {
				flowing()
				controller.enqueue(chunk)
			}
		},
		undefined,
		{ highWaterMark: 1 }
	)
	// In the order they came; `ours` for the SDK's end once this stopped it.
	const failures: { error: unknown; ours: boolean }[] = []
	let stopped = false
	const produced = produce(pipe.writable)

	produced.catch((e: unknown) => {
		failures.push({ error: e, ours: stopped })
	})

	await Promise.race([firstChunk, produced])

	const delivered = triggerSwStreamDownload(save, transferId, pipe.readable, size, () => {
		void sdkApi.cancelTransfer(transferId)
	}).catch((e: unknown) => {
		failures.push({ error: e, ours: false })
		// A browser download that died leaves the SDK writing into nothing.
		stopped = true
		void sdkApi.cancelTransfer(transferId)
	})

	const [result] = await Promise.allSettled([produced, delivered])

	const [first] = failures

	if (first !== undefined) {
		throw (failures.find(failure => !failure.ours && asErrorDTO(failure.error).kind === "Cancelled") ?? first).error
	}

	if (result.status === "rejected") {
		throw result.reason
	}

	return result.value
}
