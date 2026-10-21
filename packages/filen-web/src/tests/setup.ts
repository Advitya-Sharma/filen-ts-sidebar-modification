import { vi } from "vitest"

// Defaults every file would otherwise repeat; a test file's own vi.mock / vi.stubGlobal overrides them.
// The real client spawns a Vite `?worker`, which does not resolve under vitest.
vi.mock("@/lib/sdk/client", () => ({ sdkApi: {} }))
vi.mock("@/lib/log", () => ({ log: { debug: vi.fn(), info: vi.fn(), warn: vi.fn(), error: vi.fn(), dump: () => [] } }))

// jsdom has no ResizeObserver; an inert one lets components that merely observe mount.
if (typeof globalThis.ResizeObserver === "undefined") {
	globalThis.ResizeObserver = class {
		observe = vi.fn()
		unobserve = vi.fn()
		disconnect = vi.fn()
	}
}

// Node has no location; code that builds links from the page's own origin reads it. Defined, not
// vi.stubGlobal'd: unstubGlobals would drop it after the first test.
if (typeof globalThis.location === "undefined") {
	Object.defineProperty(globalThis, "location", { value: new URL("https://app.filen.io/"), configurable: true, writable: true })
}

// jsdom lays nothing out: a Range has no boxes (the formula's reference colours are measured from them).
if (typeof Range !== "undefined" && typeof Range.prototype.getClientRects !== "function") {
	Range.prototype.getClientRects = function () {
		return [] as unknown as DOMRectList
	}
}
