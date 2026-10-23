import { describe, expect, it, vi } from "vitest"
import { bytesSource, OpcPackage } from "@filen/office/core/opc"
import { openDocxPackage } from "@filen/office/core/docx"
import { decodeUtf8 } from "@filen/office/core/xml"
import { legacyPasswordVerifier, readProtectionHash } from "../core/docx/protection/hash"
import { verifyProtectionPassword } from "../dom/protection"
import type { DocxEvent } from "@filen/office/dom/docx"
import type { DocxConvertedDocument, DocxLegacyOpen } from "../dom/docx/legacy"
import { buildDoc, prl } from "./docFixtures"
import { workerHarness } from "./docxEncryptedFixtures"

// Word binary documents through the docx worker, as the web opens them: the worker reads the file once and holds the
// package it converted it to; the viewer's engine reads the worker's copy; Save as takes a copy to the page.

function events(host: (listener: (event: DocxEvent) => void) => { close(): void }): Promise<DocxEvent[]> {
	return new Promise(resolve => {
		const seen: DocxEvent[] = []
		const session = host(event => {
			seen.push(event)

			if (event.type === "error" || (event.type === "pages" && event.final)) {
				session.close()
				resolve(seen)
			}
		})
	})
}

function converted(opened: DocxLegacyOpen): DocxConvertedDocument {
	if (opened.kind !== "converted") throw new Error(`not converted: ${opened.kind}`)

	return opened.document
}

describe("Word binary documents in the docx worker", () => {
	it("converts once, lays the package out from the worker, and hands Save as a copy", { timeout: 30_000 }, async () => {
		const harness = workerHarness()

		try {
			const file = await buildDoc({
				pieces: [{ text: "From Word 97\r" }],
				runs: [{ start: 5, end: 9, grpprl: prl.toggle(0x0835, 1) }]
			})
			let reads = 0
			const source = bytesSource(file)
			const counted = {
				size: source.size,
				read: (offset: number, length: number) => {
					reads++

					return source.read(offset, length)
				}
			}
			const document = converted(await harness.pool.legacy.open(counted))
			const readsAfterConversion = reads

			expect(document.notConverted).toEqual([])
			await expect(document.bytes.read(0, 4)).rejects.toThrow()

			const laidOut = await events(listener => document.host(document.bytes, listener, {}))

			expect(laidOut.some(event => event.type === "opened")).toBe(true)
			expect(laidOut.some(event => event.type === "error")).toBe(false)
			// The engine read the held package, not the page's file.
			expect(reads).toBe(readsAfterConversion)

			const copy = await document.take()
			const opened = await openDocxPackage(bytesSource(copy))

			expect(opened.ok).toBe(true)
			expect(copy.length).toBe(document.bytes.size)

			document.release()
			await expect(document.take()).rejects.toThrow()
		} finally {
			harness.dispose()
		}
	})

	it("hashes an editing-restriction password only when Save as takes a copy", { timeout: 30_000 }, async () => {
		const harness = workerHarness()
		const dop = new Uint8Array(88)

		// fProtEnabled (bit 25 of the flags at 4); lKeyProtDoc at 78.
		new DataView(dop.buffer).setUint32(4, 1 << 25, true)
		new DataView(dop.buffer).setUint32(78, legacyPasswordVerifier(Uint8Array.from("Example", c => c.charCodeAt(0))), true)

		const digest = vi.spyOn(globalThis.crypto.subtle, "digest")

		try {
			const file = await buildDoc({ pieces: [{ text: "form\r" }], structures: { dop } })
			const document = converted(await harness.pool.legacy.open(bytesSource(file)))
			const laidOut = await events(listener => document.host(document.bytes, listener, {}))

			// Opening it, and laying it out, spins nothing; the password is not reported lost.
			expect(laidOut.some(event => event.type === "error")).toBe(false)
			expect(digest).not.toHaveBeenCalled()
			expect(document.notConverted).toEqual([])

			const copy = await document.take()
			const settings = decodeUtf8(await (await OpcPackage.open(bytesSource(copy))).read("word/settings.xml"))

			expect(digest.mock.calls.length).toBeGreaterThan(100_000)
			expect(settings).toMatch(/<w:documentProtection w:edit="forms" w:enforcement="1" w:cryptProviderType="rsaFull"/u)

			const opened = await openDocxPackage(bytesSource(copy))

			if (!opened.ok) throw new Error("the copy does not open")

			const hash = readProtectionHash(opened.package.settings, "documentProtection")

			if (hash === undefined || hash === "unsupported") throw new Error("no hash read back")

			digest.mockRestore()
			expect(await verifyProtectionPassword(hash, "Example")).toBe("correct")
			expect(await verifyProtectionPassword(hash, "example")).toBe("wrong")
			document.release()
		} finally {
			digest.mockRestore()
			harness.dispose()
		}
	})

	it("refuses a file it cannot convert with the reader's reason", { timeout: 30_000 }, async () => {
		const harness = workerHarness()

		try {
			const legacy = await buildDoc({ pieces: [{ text: "old\r" }], nFib: 0x65, wIdent: 0xa5dc })
			const encrypted = await buildDoc({ pieces: [{ text: "x\r" }], flags: 0x0200 | 0x0100, tablePrefix: Uint8Array.of(2, 0, 2, 0) })

			expect(await harness.pool.legacy.open(bytesSource(legacy))).toEqual({
				kind: "refused",
				refusal: { reason: "legacyVersion", nFib: 0x65 }
			})
			expect(await harness.pool.legacy.open(bytesSource(encrypted))).toEqual({
				kind: "refused",
				refusal: { reason: "encrypted", encryption: "cryptoApiRc4" }
			})
			expect(await harness.pool.legacy.open(bytesSource(new Uint8Array(32)))).toEqual({
				kind: "refused",
				refusal: { reason: "notWordBinary" }
			})
		} finally {
			harness.dispose()
		}
	})

	it("stops when aborted", async () => {
		const harness = workerHarness()

		try {
			const controller = new AbortController()

			controller.abort()

			expect(
				await harness.pool.legacy.open(bytesSource(await buildDoc({ pieces: [{ text: "x\r" }] })), { signal: controller.signal })
			).toEqual({
				kind: "aborted"
			})
		} finally {
			harness.dispose()
		}
	})
})
