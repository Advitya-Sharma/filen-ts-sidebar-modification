import { describe, expect, it } from "vitest"
import { bytesSource } from "@filen/office/core/opc"
import type { DocxEvent } from "@filen/office/dom/docx"
import { docxSealer, isSealedOfficeFile } from "@filen/office/dom/docx/editing"
import { agileStreams, decryptAgileStreams, standardStreams } from "@filen/office/testing/encryption"
import { readEncryptedContainer } from "@filen/office/core/encryption"
import { CfbFile } from "@filen/office/core/cfb"
import { docxProtectionServer } from "../dom/docx/protectionWorker"
import { buildDocx, p } from "./docxSynthetic"
import { decrypted, encryptedFile, lockedOf, workerHarness } from "./docxEncryptedFixtures"
import { cpuMs } from "./timing"

// Password-protected documents through the docx worker, as the web opens them: the spin and the decrypted package
// stay in the worker; the viewer's engine reads the worker's copy; only an editor's take() brings the package and
// the key (a structured-cloned CryptoKey) to the page.

const PLAINTEXT = buildDocx({ body: p("Held in the worker") + p("Second paragraph") })
const PASSWORD = "worker pass"
const ZIP = [0x50, 0x4b, 0x03, 0x04]

const agileFile = (spinCount = 1000): Promise<Uint8Array> => encryptedFile(agileStreams(PLAINTEXT, PASSWORD, { spinCount }))

// Any byte array in the messages the page received that starts like a zip.
function plaintextOnPage(received: readonly unknown[]): number {
	let found = 0
	const visit = (value: unknown, depth: number): void => {
		if (depth > 6 || value === null || typeof value !== "object") return
		if (value instanceof Uint8Array) {
			if (ZIP.every((byte, k) => value[k] === byte)) found++

			return
		}

		for (const item of Object.values(value)) visit(item, depth + 1)
	}

	for (const message of received) visit(message, 0)

	return found
}

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

describe("encrypted docx in the docx worker", () => {
	it(
		"prompts, refuses a wrong password, unlocks, and lays the document out without the page holding it",
		{ timeout: 30_000 },
		async () => {
			const harness = workerHarness()

			try {
				const locked = await lockedOf(await agileFile(), harness.pool.protection)

				expect([locked.scheme, locked.spinCount]).toEqual(["agile", 1000])
				expect(await locked.unlock("wrong")).toEqual({ ok: false, reason: "wrongPassword" })

				const document = decrypted(await locked.unlock(PASSWORD))

				expect(document.integrity).toBe("verified")
				expect(document.bytes.size).toBe(PLAINTEXT.length)
				await expect(document.bytes.read(0, 4)).rejects.toThrow()

				const laidOut = await events(listener => document.host(document.bytes, listener, {}))

				expect(laidOut.some(event => event.type === "opened")).toBe(true)
				expect(laidOut.some(event => event.type === "error")).toBe(false)
				// Not one decrypted byte reached the page.
				expect(plaintextOnPage(harness.received)).toBe(0)

				document.release()

				// Released: the worker holds nothing more to open or take.
				const after = await events(listener => document.host(document.bytes, listener, {}))

				expect(after.find(event => event.type === "error")).toMatchObject({ reason: "failed" })
				await expect(document.take()).rejects.toThrow()
			} finally {
				harness.dispose()
			}
		}
	)

	it("hands the editor the package and the key: the page seals with it, the worker keeps its own", { timeout: 30_000 }, async () => {
		const harness = workerHarness()

		try {
			const document = decrypted(await (await lockedOf(await agileFile(), harness.pool.protection)).unlock(PASSWORD))
			const taken = await document.take()

			expect(taken.bytes).toEqual(PLAINTEXT)
			const key = await taken.key()

			expect(key.secret).toBeInstanceOf(CryptoKey)
			expect(Reflect.get(key.secret, "extractable")).toBe(false)

			const { sealer, release } = docxSealer(taken)
			const sealed = await sealer.seal(taken.bytes)

			expect(isSealedOfficeFile(sealed)).toBe(true)

			const container = await readEncryptedContainer(await CfbFile.open(bytesSource(sealed)))
			const encryptedPackage = await container.encryptedPackage.read(0, container.encryptedPackage.size)

			expect(
				Buffer.from(
					decryptAgileStreams({ encryptionInfo: container.encryptionInfo ?? new Uint8Array(0), encryptedPackage }, PASSWORD)
				).equals(Buffer.from(PLAINTEXT))
			).toBe(true)
			release()

			// The viewer's copy in the worker is untouched by the page's release.
			expect((await document.take()).bytes).toEqual(PLAINTEXT)
			document.release()
		} finally {
			harness.dispose()
		}
	})

	it(
		"hands the editor a Standard source's package at once and derives its key in the worker at the first seal",
		{ timeout: 60_000 },
		async () => {
			const harness = workerHarness()

			try {
				const document = decrypted(
					await (
						await lockedOf(await encryptedFile(standardStreams(PLAINTEXT, PASSWORD)), harness.pool.protection)
					).unlock(PASSWORD)
				)
				const started = cpuMs()
				const taken = await document.take()

				// The copy comes without the upgrade spin (100,000 SHA-512 digests).
				expect(cpuMs() - started).toBeLessThan(250)
				expect(taken.bytes).toEqual(PLAINTEXT)

				const { sealer, release } = docxSealer(taken)

				expect(isSealedOfficeFile(await sealer.seal(taken.bytes))).toBe(true)

				const key = await taken.key()

				expect(key.template.password.spinCount).toBe(100_000)
				expect(key.template.check).toBeDefined()
				release()
				document.release()
			} finally {
				harness.dispose()
			}
		}
	)

	it("drops the worker's copy once the editor has its own: the key stays for the first seal", { timeout: 30_000 }, async () => {
		const harness = workerHarness()

		try {
			const document = decrypted(await (await lockedOf(await agileFile(), harness.pool.protection)).unlock(PASSWORD))
			const taken = await document.take()

			document.dropPackage()

			const after = await events(listener => document.host(document.bytes, listener, {}))

			expect(after.find(event => event.type === "error")).toMatchObject({ reason: "failed" })
			await expect(document.take()).rejects.toThrow()

			const { sealer, release } = docxSealer(taken)

			expect(isSealedOfficeFile(await sealer.seal(taken.bytes))).toBe(true)
			release()
			document.release()
		} finally {
			harness.dispose()
		}
	})

	it("stops an unlock in the worker at the page's AbortSignal, and unlocks on a retry", { timeout: 30_000 }, async () => {
		const harness = workerHarness()

		try {
			const locked = await lockedOf(await agileFile(300_000), harness.pool.protection)
			const controller = new AbortController()
			const started = performance.now()
			const pending = locked.unlock(PASSWORD, { signal: controller.signal })

			setTimeout(() => controller.abort(), 20)
			expect(await pending).toEqual({ ok: false, reason: "aborted" })
			expect(performance.now() - started).toBeLessThan(1500)

			const aborted = new AbortController()

			aborted.abort()
			expect(await harness.pool.protection.open(bytesSource(await agileFile()), { signal: aborted.signal })).toEqual({
				kind: "aborted"
			})
			locked.dispose()
		} finally {
			harness.dispose()
		}
	})

	it("says plain for a zip, and keeps the idle worker alive while a document is held", { timeout: 30_000 }, async () => {
		const harness = workerHarness(20)

		try {
			expect(await harness.pool.protection.open(bytesSource(PLAINTEXT))).toEqual({ kind: "plain" })

			const document = decrypted(await (await lockedOf(await agileFile(), harness.pool.protection)).unlock(PASSWORD))

			await new Promise(resolve => setTimeout(resolve, 100))
			// Held: the idle worker was not ended.
			expect(harness.terminated()).toBe(0)
			document.release()
			await new Promise(resolve => setTimeout(resolve, 100))
			expect(harness.terminated()).toBe(1)
		} finally {
			harness.dispose()
		}
	})

	it("never answers ok for an unlock the page disposed of meanwhile: no plaintext reaches the page", { timeout: 60_000 }, async () => {
		const harness = workerHarness()

		try {
			const locked = await lockedOf(await agileFile(200_000), harness.pool.protection)
			const unlocking = locked.unlock(PASSWORD).then(
				answer => ({ answer }),
				(error: unknown) => ({ error: String(error) })
			)

			await new Promise(resolve => setTimeout(resolve, 30))
			locked.dispose()

			const outcome = await unlocking

			expect("answer" in outcome && outcome.answer.ok).toBe(false)
			await new Promise(resolve => setTimeout(resolve, 300))
			expect(plaintextOnPage(harness.received)).toBe(0)
		} finally {
			harness.dispose()
		}
	})

	it("hands out no key once released while a Standard key's upgrade runs", { timeout: 60_000 }, async () => {
		const harness = workerHarness()

		try {
			const document = decrypted(
				await (await lockedOf(await encryptedFile(standardStreams(PLAINTEXT, PASSWORD)), harness.pool.protection)).unlock(PASSWORD)
			)
			const taken = await document.take()
			const keying = taken.key().then(
				key => ({ key }),
				(error: unknown) => ({ error: String(error) })
			)

			await new Promise(resolve => setTimeout(resolve, 20))
			document.release()
			expect("error" in (await keying)).toBe(true)
		} finally {
			harness.dispose()
		}
	})

	it("a closed document answers no take or key, also when closed in the same turn", async () => {
		const server = docxProtectionServer()
		const file = await agileFile(10)

		expect((await server.protect(1, bytesSource(file))).kind).toBe("needsPassword")
		expect((await server.unlock(1, 1, PASSWORD)).ok).toBe(true)

		const keying = server.key(1, 2).then(
			key => ({ key }),
			(error: unknown) => ({ error: String(error) })
		)

		expect(server.close(1)).toBe(true)
		expect("error" in (await keying)).toBe(true)
		expect(() => server.take(1)).toThrow()
	})
})
