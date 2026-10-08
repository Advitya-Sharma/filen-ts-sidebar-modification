import { describe, expect, it } from "vitest"
import { FakeApp } from "../fakeApp"

function ok(result: ReturnType<FakeApp["execute"]>): Record<string, unknown> {
	if (!result.ok) {
		throw new Error(result.error)
	}

	return result.result as Record<string, unknown>
}

describe("FakeApp", () => {
	it("finds items by name tokens, folding plurals", () => {
		const app = new FakeApp()
		const found = ok(
			app.execute({
				name: "search_items",
				args: {
					query: "invoices 2024"
				}
			})
		)["results"] as {
			handle: string
		}[]

		expect(found.map(item => item.handle).sort()).toEqual(["f1", "f2", "f3"])
		expect(app.exposed.has("f1")).toBe(true)
	})

	it("counts what a search left out from the pool its results come from, like the app", () => {
		const search = (query: string) =>
			ok(
				new FakeApp().execute({
					name: "search_items",
					args: {
						query
					}
				})
			) as {
				results: { handle: string }[]
				more?: number
			}

		// One name matches both tokens; the partial matches it outranks are not left-out results.
		expect(search("PDF files").results.map(item => item.handle)).toEqual(["f7"])
		expect(search("PDF files").more).toBeUndefined()
		expect(search("pdf").results).toHaveLength(12)
		expect(search("pdf").more).toBe(2)
	})

	it("refuses handles the conversation has not seen and offers candidates for names", () => {
		const app = new FakeApp()
		const unseen = app.execute({
			name: "read_note",
			args: {
				handle: "n3"
			}
		})
		const named = app.execute({
			name: "read_note",
			args: {
				handle: "Groceries"
			}
		})

		expect(unseen.ok).toBe(false)
		expect(!unseen.ok && unseen.invalid).toBe(true)
		expect(!named.ok && named.candidates).toEqual([
			{
				handle: "n1",
				name: "Groceries"
			}
		])
		expect(app.exposed.has("n1")).toBe(true)
	})

	it("moves items and keeps sibling names unique case-insensitively", () => {
		const app = new FakeApp()

		app.exposed.add("f1").add("f26").add("d9").add("d8")

		expect(
			ok(
				app.execute({
					name: "move_items",
					args: {
						handles: ["f1"],
						destination: "d8"
					}
				})
			)
		).toEqual({
			moved: ["Invoice_2024-01_Telekom.pdf"],
			to: "/Archive"
		})

		const clash = app.execute({
			name: "create_directory",
			args: {
				parent: "d8",
				name: "2023"
			}
		})

		expect(clash.ok).toBe(false)
		expect(!clash.ok && clash.invalid).toBe(false)
	})

	it("cannot put a directory inside itself", () => {
		const app = new FakeApp()

		app.exposed.add("d8").add("d9")

		expect(
			app.execute({
				name: "move_items",
				args: {
					handles: ["d8"],
					destination: "d9"
				}
			}).ok
		).toBe(false)
	})

	it("resolves paths given where a handle belongs", () => {
		const app = new FakeApp()

		expect(
			app.resolvePaths({
				destination: "/Archive/",
				handles: ["/Downloads/report_Q3.pdf", "f1"],
				name: "/not/a/handle"
			})
		).toEqual({
			destination: "d8",
			handles: ["f33", "f1"],
			name: "/not/a/handle"
		})
	})

	it("lists a directory by path even when a handle is given too", () => {
		const app = new FakeApp()
		const listing = ok(
			app.execute({
				name: "list_directory",
				args: {
					handle: "d0",
					path: "/Documents/Invoices"
				}
			})
		)

		expect(listing["directory"]).toEqual({
			handle: "d2",
			path: "/Documents/Invoices"
		})
	})

	it("plays a playlist from its first track and controls playback", () => {
		const app = new FakeApp()

		app.exposed.add("p3")
		ok(
			app.execute({
				name: "play",
				args: {
					handle: "p3"
				}
			})
		)
		ok(
			app.execute({
				name: "audio_control",
				args: {
					action: "next"
				}
			})
		)

		expect(app.state.context.playing).toMatchObject({
			handle: "f18",
			playlist: "p3"
		})
	})

	it("trashes drive items and notes but never the root", () => {
		const app = new FakeApp()

		app.exposed.add("n15").add("f32")
		ok(
			app.execute({
				name: "trash_items",
				args: {
					handles: ["n15", "f32"]
				}
			})
		)

		expect(app.state.notes.find(note => note.handle === "n15")?.trashed).toBe(true)
		expect(
			app.execute({
				name: "trash_items",
				args: {
					handles: ["d0"]
				}
			}).ok
		).toBe(false)
	})

	it("describes the context with handles", () => {
		const app = new FakeApp({
			directory: "d2",
			selection: ["f3"]
		})

		expect(app.describeContext()).toBe(
			"screen: drive; current directory: /Documents/Invoices (handle d2); selected: Invoice_2024-11_Hetzner.pdf (handle f3); nothing playing"
		)
		expect([...app.exposed].sort()).toEqual(["d0", "d2", "f3"])
	})
})
