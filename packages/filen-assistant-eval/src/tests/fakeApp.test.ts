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
	it("filters by dates resolved in code and says which days it kept", () => {
		const app = new FakeApp()
		const listing = ok(
			app.execute({
				name: "list_directory",
				args: {
					path: "/Downloads",
					period: "this_month"
				}
			})
		) as {
			items: { handle: string; modified: string }[]
			dates: string
		}

		expect(listing.items).toEqual([
			expect.objectContaining({ handle: "f33", modified: "2025-10-01" }),
			expect.objectContaining({ handle: "f36", modified: "2025-10-02" })
		])
		expect(listing.dates).toBe("modified 2025-10-01 to 2025-10-31")

		const explicit = ok(app.execute({ name: "search_items", args: { query: "invoice", from: "2024", to: "2024-06" } })) as {
			results: { handle: string }[]
		}

		expect(explicit.results.map(item => item.handle).sort()).toEqual(["f1", "f2"])
		expect(app.execute({ name: "search_items", args: { query: "invoice", from: "last year" } })).toMatchObject({
			ok: false,
			invalid: true
		})
		expect(app.execute({ name: "search_items", args: { query: " ", period: "this_month" } })).toMatchObject({
			ok: false,
			invalid: true
		})
	})

	it("ranks directories by total size with files, showing what it sorted by", () => {
		const listing = ok(new FakeApp().execute({ name: "list_directory", args: { path: "/Documents", sort: "largest" } })) as {
			items: { handle: string; size?: string }[]
			note?: string
		}

		expect(listing.items.map(item => item.handle)).toEqual(["d3", "f10", "d2", "f8", "f44", "f9"])
		expect(listing.items[0]?.size).toBe("2.3 MiB")
		expect(listing.note).toBeUndefined()

		const root = ok(new FakeApp().execute({ name: "list_directory", args: { path: "/", sort: "largest", kind: "directory" } })) as {
			items: { handle: string }[]
		}

		expect(root.items[0]?.handle).toBe("d4")
	})

	it("filters notes by their edited date before the cut, archived ones included", () => {
		const app = new FakeApp()

		for (let index = 0; index < 45; index++) {
			app.state.notes.push({
				handle: `n${String(100 + index)}`,
				title: `Recent ${String(index)}`,
				type: "text",
				content: "",
				tags: [],
				pinned: false,
				favorite: false,
				archived: false,
				trashed: false,
				edited: "2025-10-07T10:00:00Z"
			})
		}

		const old = app.state.notes.find(note => note.handle === "n2")

		if (old !== undefined) {
			old.archived = true
			old.edited = "2025-03-15T10:00:00Z"
		}

		const all = ok(app.execute({ name: "list_notes", args: {} })) as { notes: { handle: string }[]; more?: number }

		expect(all.notes).toHaveLength(40)
		expect(all.more).toBeGreaterThan(0)
		expect(all.notes.map(note => note.handle)).not.toContain("n2")

		const march = ok(app.execute({ name: "list_notes", args: { period: "custom", from: "2025-03", to: "2025-03" } })) as {
			notes: { handle: string }[]
			dates: string
		}

		expect(march.notes.map(note => note.handle)).toEqual(["n2"])
		expect(march.dates).toBe("edited 2025-03-01 to 2025-03-31")
		expect(app.execute({ name: "list_notes", args: { query: "einkaufsliste", period: "today" } })).toMatchObject({
			result: { notes: [{ handle: "n2" }], note: expect.stringContaining("None of these is dated 2025-10-08") }
		})
	})

	it("lists the Photos timeline by the date taken", () => {
		const photos = ok(new FakeApp().execute({ name: "list_photos", args: { period: "custom", from: "2024-05", to: "2024-05" } })) as {
			photos: { handle: string; taken: string }[]
		}

		expect(photos.photos.map(photo => photo.handle)).toEqual(["f13", "f12", "f11"])
	})

	it("lists recent uploads by upload date and notes how far back they reach", () => {
		const recent = ok(new FakeApp().execute({ name: "recent_items", args: { kind: "image" } })) as {
			items: { handle: string; uploaded: string }[]
		}

		expect(recent.items).toEqual([expect.objectContaining({ handle: "f14", uploaded: "2025-10-06" })])

		const lastYear = ok(new FakeApp().execute({ name: "recent_items", args: { period: "last_year" } })) as {
			items: unknown[]
			note?: string
		}

		expect(lastYear.items).toEqual([])
		expect(lastYear.note).toContain("Recents reaches back to uploads of 2025-08-30")
	})

	it("reports directory totals and storage like the app", () => {
		const app = new FakeApp({ directory: "d4" })

		expect(ok(app.execute({ name: "get_item_info", args: { handle: "d4" } }))["size"]).toBe("195.6 MiB in 5 files and 2 directories")
		expect(ok(app.execute({ name: "get_storage", args: {} }))).toEqual({
			used: "3.2 GiB",
			total: "10 GiB",
			free: "6.8 GiB",
			usedPercent: 32,
			versions: "450 MiB"
		})
	})

	it("lists the drive's own lists, with shared items outside My Drive", () => {
		const app = new FakeApp()
		const section = (name: string) =>
			(ok(app.execute({ name: "list_section", args: { section: name } })) as { items: { handle: string }[] }).items.map(
				item => item.handle
			)

		expect(section("favorites")).toEqual(["d5", "f8"])
		expect(section("trash")).toEqual(["d49", "f84", "f45"])
		expect(section("shared_with_me")).toEqual(["d51", "d50", "f83"])
		expect(section("shared_by_me")).toEqual(["d11", "f10"])
		expect(section("public_links")).toEqual(["d5", "f33"])

		// Neither search nor a path reaches into someone else's drive.
		expect(ok(app.execute({ name: "search_items", args: { query: "Lohnsteuerbescheinigung" } }))["results"]).toEqual([])
		expect(app.execute({ name: "list_directory", args: { path: "/Steuerunterlagen 2025" } }).ok).toBe(false)
		expect(ok(app.execute({ name: "list_directory", args: { handle: "d50" } }))["directory"]).toEqual({
			handle: "d50",
			sharedBy: "anna.schmidt@example.com"
		})
	})

	it("reports links and shares, also through a parent directory", () => {
		const app = new FakeApp({ directory: "d11", selection: ["f27"] })

		expect(ok(app.execute({ name: "get_sharing", args: { handle: "f27" } }))).toEqual({
			handle: "f27",
			name: "roadmap.md",
			publicLink: "none",
			sharedWith: [],
			sharedThrough: { directory: "Website relaunch", with: ["lena.hoffmann@example.com", "tom.becker@example.com"] }
		})
		expect(ok(app.execute({ name: "get_sharing", args: { handle: "d11" } }))["sharedWith"]).toEqual([
			"lena.hoffmann@example.com",
			"tom.becker@example.com"
		])
	})

	it("finds contacts with the chat to write them in, and reports it as returned", () => {
		const app = new FakeApp()
		const found = app.execute({ name: "find_contacts", args: { query: "carlos.ruiz@example.com" } })

		expect(found).toEqual({
			ok: true,
			result: { contacts: [{ name: "Carlos Ruiz", email: "carlos.ruiz@example.com", chat: "c3" }] },
			returned: [{ handle: "c3", name: "Carlos Ruiz" }]
		})
		expect(app.exposed.has("c3")).toBe(true)
		expect(ok(app.execute({ name: "find_contacts", args: { query: "Lena" } }))["note"]).toContain("no chat")
	})

	it("matches contacts only on an exact address or every word, so a domain grounds no one's chat", () => {
		const app = new FakeApp()

		for (const query of ["example.com", "carlos@example.com", "Carlos Becker"]) {
			expect(app.execute({ name: "find_contacts", args: { query } })).toEqual({
				ok: true,
				result: { contacts: [], note: "No contact matches. Ask the user for the contact's name or exact email address." },
				returned: []
			})
		}
	})

	it("narrows a drive list by name or by whom an item is shared with, and leaves an owned item's sharing unchecked", () => {
		const app = new FakeApp({ directory: "d11", selection: ["f27"] })

		expect(
			(ok(app.execute({ name: "list_section", args: { section: "trash", query: "bewerbung" } })) as { items: { handle: string }[] })
				.items
		).toMatchObject([{ handle: "f84" }])
		expect(
			(
				ok(app.execute({ name: "list_section", args: { section: "shared_with_me", query: "anna" } })) as {
					items: { handle: string }[]
				}
			).items
		).toMatchObject([{ handle: "d50" }])
		expect(ok(app.execute({ name: "get_item_info", args: { handle: "f27" } }))["sharing"]).toBe(
			"unknown: this tool does not show shares or links"
		)
	})

	it("reports transfers with progress, failures and the item an upload became", () => {
		const transfers = ok(new FakeApp().execute({ name: "get_transfers", args: {} }))

		expect(transfers).toMatchObject({ running: 2, overall: "23%", failed: 1 })
		expect(transfers["transfers"]).toContainEqual({ name: "report_Q3.pdf", handle: "f33", kind: "upload", status: "done" })
	})
})
