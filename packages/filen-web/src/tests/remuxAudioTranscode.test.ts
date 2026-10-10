import { describe, expect, it, vi } from "vitest"
import type { RemuxStats } from "@/lib/media/remuxTypes"
import { TranscodeWriter, type SampleItem } from "@/workers/remux/audioTranscode"
import type { EncoderTarget } from "@/workers/remux/audioMix.logic"
import type { FragmentedTrackOutput } from "@/workers/remux/output"

// An encoder that fails as soon as it is configured, the way a platform encoder reports a broken config.
class FailingEncoder {
	public state = "configured"
	public encodeQueueSize = 0
	private readonly onError: (error: Error) => void

	public constructor(init: { error: (error: Error) => void }) {
		this.onError = init.error
	}

	public configure(): void {
		this.onError(new Error("encoder broke"))
	}

	public close(): void {
		this.state = "closed"
	}
}

describe("TranscodeWriter", () => {
	it("closes the decoded sample when the encoder has already failed", async () => {
		vi.stubGlobal("AudioEncoder", FailingEncoder)

		const writer = new TranscodeWriter({
			target: { target: "aac", codec: "mp4a.40.2" } satisfies EncoderTarget,
			output: {} as FragmentedTrackOutput,
			stats: { transcodeSeconds: 0, transcodedMediaSeconds: 0 } as RemuxStats,
			layout: () => []
		})
		const close = vi.fn()
		const item = { sample: { close }, timestamp: 0, duration: 0.02, type: "key" } as unknown as SampleItem

		await expect(writer.write(item)).rejects.toThrow("encoder broke")
		expect(close).toHaveBeenCalledOnce()
	})
})
