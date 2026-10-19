import { describe, expect, it } from "vitest"
import { ZipArchive, ZipError } from "@filen/office/core/opc"
import { buildZip } from "../zipBuilder"
import { countingSource } from "./generate"

// The end-record search looks back through up to a mebibyte for an end record that leads to a directory:
// a file full of decoys costs a bounded number of reads (each one a native call on mobile).

// `size` bytes holding an end record every 42 bytes, each after a ZIP64 locator pointing somewhere else.
function decoys(size: number): Uint8Array {
	const bytes = new Uint8Array(size)
	const view = new DataView(bytes.buffer)

	for (let at = size - 22; at >= 42; at -= 42) {
		view.setUint32(at, 0x06054b50, true)
		view.setUint32(at + 16, 999_999, true)
		view.setUint32(at - 20, 0x07064b50, true)
		view.setUint32(at - 12, (at * 7919) % (size - 64), true)
	}

	return bytes
}

describe("end record search", () => {
	it("gives up on a tail of decoy end records after a bounded number of reads", async () => {
		const source = countingSource(decoys(2 * 1024 * 1024))

		await expect(ZipArchive.open(source)).rejects.toSatisfy((error: unknown) => error instanceof ZipError && error.code === "notZip")
		expect(source.reads.length).toBeLessThan(1000)
	})

	it("still finds an archive behind a comment holding a few decoys", async () => {
		// End records whose directory lies past the file.
		const decoy =
			"PK\u0005\u0006\u0000\u0000\u0000\u0000\u0001\u0000\u0001\u0000\u002e\u0000\u0000\u0000\u007f\u007f\u007f\u007f\u0000\u0000"
		const bytes = buildZip([{ name: "a.xml", data: new TextEncoder().encode("<a/>") }], { comment: decoy.repeat(8) })
		const archive = await ZipArchive.open(countingSource(bytes))

		expect(archive.entries.map(entry => entry.name)).toEqual(["a.xml"])
	})
})
