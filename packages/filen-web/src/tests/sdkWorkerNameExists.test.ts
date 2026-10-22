import { beforeEach, describe, expect, it, vi } from "vitest"
import type { Remote } from "comlink"
import type { Dir, File, NormalDirsAndFiles, StringifiedClient } from "@filen/sdk-rs"
import type { SdkWorkerApi } from "@/workers/sdk.worker"

// A fresh file's name check, the guard against uploading onto (and so versioning) a file already there.

const { exposed, fakeClient } = vi.hoisted(() => ({
	exposed: new Map<"api", unknown>(),
	fakeClient: {
		root: () => ({ uuid: "root" }),
		listDir: vi.fn<(dir: unknown) => Promise<NormalDirsAndFiles>>(),
		findItemInDir: vi.fn(),
		free: vi.fn()
	}
}))

vi.mock("@filen/sdk-rs", () => ({
	default: vi.fn(),
	initThreadPool: vi.fn(),
	UnauthClient: { from_config: () => ({ free: vi.fn(), fromStringified: () => fakeClient }) }
}))

vi.mock("comlink", () => ({
	expose: (api: unknown) => {
		exposed.set("api", api)
	},
	proxy: (value: unknown) => value,
	transfer: (value: unknown) => value
}))

await import("@/workers/sdk.worker")

const api = exposed.get("api") as Remote<SdkWorkerApi>

function file(name: string): File {
	return { uuid: `f-${name}`, meta: { type: "decoded", data: { name } } } as unknown as File
}

function dir(name: string): Dir {
	return { uuid: `d-${name}`, meta: { type: "decoded", data: { name } } } as unknown as Dir
}

beforeEach(async () => {
	await api.injectClient("session" as unknown as StringifiedClient)
	fakeClient.listDir.mockReset()
})

describe("nameExistsInDirectory", () => {
	it("detects a file whose name differs only in letter case, as the server would version it", async () => {
		fakeClient.listDir.mockResolvedValue({ dirs: [], files: [file("Report.DOCX")] })

		await expect(api.nameExistsInDirectory(null, "report.docx")).resolves.toBe(true)
		await expect(api.nameExistsInDirectory(null, "İstanbul.txt")).resolves.toBe(false)
		expect(fakeClient.findItemInDir).not.toHaveBeenCalled()
	})

	it("detects a directory of that name, in any letter case", async () => {
		fakeClient.listDir.mockResolvedValue({ dirs: [dir("ÄRGER")], files: [] })

		await expect(api.nameExistsInDirectory(null, "ärger")).resolves.toBe(true)
	})

	it("is false for a free name", async () => {
		fakeClient.listDir.mockResolvedValue({ dirs: [dir("Docs")], files: [file("notes.txt")] })

		await expect(api.nameExistsInDirectory(null, "notes 2.txt")).resolves.toBe(false)
	})
})
