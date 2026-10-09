import { describe, expect, it } from "vitest"
import { outputVideoCodecString } from "@/workers/remux/codecs.logic"

describe("outputVideoCodecString", () => {
	it("names HEVC by the hvc1 sample entry the muxer writes, with or without a configuration record", () => {
		expect(outputVideoCodecString("hevc", { codec: "hev1.1.6.L93.B0" })).toBe("hvc1.1.6.L93.B0")
		expect(outputVideoCodecString("hevc", { codec: "hvc1.2.4.L120.90" })).toBe("hvc1.2.4.L120.90")
	})

	it("leaves other codecs as the decoder config says", () => {
		expect(outputVideoCodecString("avc", { codec: "avc1.64001f" })).toBe("avc1.64001f")
		expect(outputVideoCodecString("vp9", { codec: "vp09.00.31.08" })).toBe("vp09.00.31.08")
	})
})
