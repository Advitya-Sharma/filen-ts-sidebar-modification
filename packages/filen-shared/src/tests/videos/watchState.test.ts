import { describe, expect, it } from "vitest"
import {
	compactWatchState,
	compareStamps,
	crossesCompletion,
	deleteRecord,
	deriveWatchState,
	emptyWatchState,
	hideFromContinueWatching,
	hlcObserve,
	hlcPhysicalMs,
	hlcTick,
	isCompletePosition,
	isResumablePosition,
	itemKey,
	markUnwatched,
	markWatched,
	MAX_CLOCK_SKEW_MS,
	mergeWatchState,
	parseWatchState,
	parseWatchStateText,
	reconcileWatchState,
	seriesKey,
	STALE_PROGRESS_MS,
	TOMBSTONE_TTL_MS,
	utf8Length,
	WATCH_STATE_HARD_LIMIT_BYTES,
	withAudioLanguage,
	withCompletion,
	withProgress,
	withSubtitleLanguage,
	withSubtitleOffset,
	type ItemRec,
	type SeriesPrefs,
	type Stamp,
	type WatchStateDoc
} from "@filen/shared/videos"

const NOW = Date.UTC(2026, 9, 9)
const DAY = 86_400_000

function stamp(t: number, d = "dev-a"): Stamp {
	return {
		t,
		d
	}
}

function progress({
	t,
	pos,
	dur = 3600,
	lp = NOW,
	u = "file-1",
	sz = 1000,
	d = "dev-a"
}: Partial<Record<"t" | "pos" | "dur" | "lp" | "sz", number>> & {
	t: number
	pos: number
	u?: string
	d?: string
}): ItemRec {
	return withProgress({
		rec: undefined,
		stamp: stamp(t, d),
		pos,
		dur,
		lp,
		uuid: u,
		size: sz
	})
}

function docWith(items: Record<string, ItemRec>, extra: Partial<WatchStateDoc> = {}): WatchStateDoc {
	return {
		...emptyWatchState(),
		...extra,
		items
	}
}

function parsed(value: unknown) {
	const result = parseWatchState(value)

	if (!result.ok) {
		throw new Error(`parse failed: ${result.reason}`)
	}

	return result
}

// Small seeded PRNG (mulberry32) so property runs are reproducible.
function prng(seed: number): () => number {
	let state = seed >>> 0

	return () => {
		state = (state + 0x6d2b79f5) >>> 0

		let t = state

		t = Math.imul(t ^ (t >>> 15), t | 1)
		t ^= t + Math.imul(t ^ (t >>> 7), t | 61)

		return ((t ^ (t >>> 14)) >>> 0) / 4294967296
	}
}

describe("hybrid logical clock", () => {
	it("jumps to the wall clock when it is ahead", () => {
		expect(
			hlcTick({
				clock: 5,
				now: 1000
			})
		).toBe(1_000_000)
	})

	it("counts within one millisecond and stays monotonic past the counter span", () => {
		let clock = hlcTick({
			clock: 0,
			now: 1000
		})

		for (let i = 0; i < 1500; i++) {
			const next = hlcTick({
				clock,
				now: 1000
			})

			expect(next).toBeGreaterThan(clock)
			clock = next
		}

		expect(hlcPhysicalMs(clock)).toBe(1001)
	})

	it("keeps counting when the wall clock goes backwards", () => {
		expect(
			hlcTick({
				clock: 2_000_005,
				now: 1000
			})
		).toBe(2_000_006)
	})

	it("observes remote stamps so the next local stamp is newer", () => {
		const seen = (NOW + 60_000) * 1000 + 7
		const clock = hlcObserve({
			clock: NOW * 1000,
			seen,
			now: NOW
		})

		expect(clock).toBe(seen)
		expect(
			hlcTick({
				clock,
				now: NOW
			})
		).toBeGreaterThan(seen)
	})

	it("follows observed stamps at most a year ahead", () => {
		expect(
			hlcObserve({
				clock: 0,
				seen: (NOW + MAX_CLOCK_SKEW_MS + DAY) * 1000,
				now: NOW
			})
		).toBe((NOW + MAX_CLOCK_SKEW_MS) * 1000)
	})

	it("ignores a broken wall clock", () => {
		expect(
			hlcTick({
				clock: 41,
				now: Number.NaN
			})
		).toBe(42)
	})
})

describe("parseWatchState", () => {
	it("round-trips a document and preserves unknown fields at every level", () => {
		const doc = {
			v: 1,
			gc: 123,
			future: {
				x: 1
			},
			items: {
				"s:a": {
					p: {
						t: 5,
						d: "dev",
						pos: 1,
						dur: 2,
						lp: 3,
						u: "u",
						sz: 4,
						part: 1,
						extra: "keep"
					},
					w: {
						t: 6,
						d: "dev",
						pl: true,
						pc: 2
					},
					o: {
						t: 7,
						d: "dev",
						subMs: -250
					},
					h: {
						t: 8,
						d: "dev",
						v: true
					},
					z: [1, 2]
				},
				"s:gone": {
					x: {
						t: 4,
						d: "dev",
						e: 10
					}
				}
			},
			series: {
				"d:s": {
					al: {
						t: 9,
						d: "dev",
						v: "eng"
					},
					sl: {
						t: 3,
						d: "dev",
						v: null,
						more: true
					},
					more: true
				}
			}
		}

		const result = parsed(JSON.parse(JSON.stringify(doc)))

		expect(result.doc).toEqual(doc)
		expect(result.dropped).toBe(0)
		expect(result.maxStamp).toBe(9)
	})

	it("drops malformed registers, records and keys individually", () => {
		const result = parsed({
			v: 1,
			gc: "x",
			items: {
				"s:good": {
					p: {
						t: 1,
						d: "dev",
						pos: -1,
						dur: 2,
						lp: 3,
						u: "u",
						sz: 4
					},
					w: {
						t: 2,
						d: "dev",
						pl: true,
						pc: 1
					}
				},
				"s:bad": "nope",
				"s:allBad": {
					w: {
						t: 1,
						d: "",
						pl: true,
						pc: 1
					}
				},
				"x:wrongPrefix": {},
				"s:": {}
			},
			series: {
				"d:s": {
					al: {
						t: 1,
						d: "dev",
						v: 5
					}
				},
				"d:ok": {
					sl: {
						t: 3,
						d: "dev",
						v: "eng"
					}
				}
			}
		})

		expect(result.doc.items).toEqual({
			"s:good": {
				w: {
					t: 2,
					d: "dev",
					pl: true,
					pc: 1
				}
			}
		})
		expect(result.doc.gc).toBeUndefined()
		expect(Object.keys(result.doc.series)).toEqual(["d:ok"])
		expect(result.dropped).toBe(7)
		expect(result.maxStamp).toBe(3)
	})

	it("keeps registers stamped in the future so every replica reads the same document", () => {
		const future = (NOW + MAX_CLOCK_SKEW_MS + DAY) * 1000
		const doc = {
			v: 1,
			items: {
				"s:a": {
					w: {
						t: future,
						d: "dev",
						pl: true,
						pc: 1
					}
				}
			},
			series: {}
		}
		const result = parsed(doc)

		expect(result.doc).toEqual(doc)
		expect(result.maxStamp).toBe(future)
	})

	it("applies tombstones while parsing", () => {
		const result = parsed({
			v: 1,
			items: {
				"s:a": {
					p: {
						t: 4,
						d: "dev",
						pos: 1,
						dur: 2,
						lp: 3,
						u: "u",
						sz: 4
					},
					w: {
						t: 6,
						d: "dev",
						pl: true,
						pc: 1
					},
					x: {
						t: 5,
						d: "dev",
						e: 0
					}
				}
			}
		})

		expect(Object.keys(result.doc.items["s:a"] ?? {})).toEqual(["w", "x"])
	})

	it("reports non-documents and newer schema versions without throwing", () => {
		for (const value of [null, 1, "x", [], {}, { v: 0 }, { v: 1, items: [] }, { v: 1, series: 5 }]) {
			expect(parseWatchState(value)).toEqual({
				ok: false,
				reason: "malformed"
			})
		}

		expect(
			parseWatchState({
				v: 2,
				items: {}
			})
		).toEqual({
			ok: false,
			reason: "unsupportedVersion"
		})
		expect(parseWatchStateText("{not json")).toEqual({
			ok: false,
			reason: "malformed"
		})
	})

	it("treats missing maps as empty", () => {
		expect(parsed({ v: 1 }).doc).toEqual(emptyWatchState())
	})

	it("never lets a __proto__ key reach an object prototype", () => {
		const withOwnProto = (target: Record<string, unknown>): Record<string, unknown> =>
			Object.defineProperty(target, "__proto__", {
				value: {
					polluted: 1
				},
				enumerable: true
			})
		const text = JSON.stringify(
			withOwnProto({
				v: 1,
				items: {
					"s:a": withOwnProto({
						h: {
							t: 1,
							d: "x",
							v: true
						}
					})
				},
				series: {}
			})
		)

		expect(text).toContain("__proto__")

		const result = parseWatchStateText(text)

		expect(result.ok).toBe(true)

		if (!result.ok) {
			return
		}

		const rec = result.doc.items["s:a"]

		expect(Object.getPrototypeOf(result.doc)).toBe(Object.prototype)
		expect(rec === undefined ? null : Object.getPrototypeOf(rec)).toBe(Object.prototype)
		expect(Object.hasOwn(result.doc, "__proto__")).toBe(false)
		expect(({} as Record<string, unknown>)["polluted"]).toBeUndefined()
	})

	it("survives random garbage", () => {
		const random = prng(7)
		const pick = <T>(values: readonly T[]): T => values[Math.floor(random() * values.length)] as T
		const garbage = (depth: number): unknown => {
			const kind = Math.floor(random() * (depth > 3 ? 5 : 7))

			switch (kind) {
				case 0:
					return null
				case 1:
					return random() * 1e6 - 1e5
				case 2:
					return pick(["", "x", "s:a", "dev"])
				case 3:
					return random() > 0.5
				case 4:
					return Math.floor(random() * 1e9)
				case 5:
					return Array.from(
						{
							length: 3
						},
						() => garbage(depth + 1)
					)
				default: {
					const out: Record<string, unknown> = {}

					for (let i = 0; i < 4; i++) {
						out[pick(["v", "items", "series", "s:a", "d:b", "p", "w", "o", "h", "t", "d", "pos", "pl", "pc", "x"])] = garbage(
							depth + 1
						)
					}

					return out
				}
			}
		}

		for (let i = 0; i < 2000; i++) {
			const value = random() > 0.5 ? { v: 1, items: garbage(1), series: garbage(1) } : garbage(0)

			expect(() => parseWatchState(value)).not.toThrow()
		}
	})
})

describe("mergeWatchState", () => {
	it("takes the register with the higher stamp, then the higher device id", () => {
		const local = docWith({
			"s:a": progress({
				t: 10,
				pos: 100
			}),
			"s:b": progress({
				t: 10,
				pos: 100,
				d: "dev-a"
			})
		})
		const remote = docWith({
			"s:a": progress({
				t: 11,
				pos: 200
			}),
			"s:b": progress({
				t: 10,
				pos: 300,
				d: "dev-b"
			})
		})
		const { doc } = mergeWatchState({
			local,
			remote
		})

		expect(doc.items["s:a"]?.p?.pos).toBe(200)
		expect(doc.items["s:b"]?.p?.pos).toBe(300)
	})

	it("merges registers of one record independently", () => {
		const local = docWith({
			"s:a": withSubtitleOffset({
				rec: progress({
					t: 20,
					pos: 600
				}),
				stamp: stamp(21),
				subMs: 500
			})
		})
		const remote = docWith({
			"s:a": withCompletion({
				rec: undefined,
				stamp: stamp(30, "dev-b")
			})
		})
		const { doc, localContributed, remoteContributed } = mergeWatchState({
			local,
			remote
		})

		expect(doc.items["s:a"]).toEqual({
			p: local.items["s:a"]?.p,
			o: local.items["s:a"]?.o,
			w: remote.items["s:a"]?.w
		})
		expect(localContributed).toBe(true)
		expect(remoteContributed).toBe(true)
	})

	it("returns a side unchanged when the other adds nothing", () => {
		const local = docWith(
			{
				"s:a": progress({
					t: 50,
					pos: 600
				})
			},
			{
				gc: 5
			}
		)
		const older = docWith({
			"s:a": progress({
				t: 40,
				pos: 100
			})
		})

		const first = mergeWatchState({
			local,
			remote: older
		})

		expect(first.doc).toBe(local)
		expect(first.localContributed).toBe(true)
		expect(first.remoteContributed).toBe(false)

		const second = mergeWatchState({
			local: older,
			remote: local
		})

		expect(second.doc).toBe(local)
		expect(second.localContributed).toBe(false)
		expect(second.remoteContributed).toBe(true)

		const same = mergeWatchState({
			local,
			remote: structuredClone(local)
		})

		expect(same.doc).toBe(local)
		expect(same.localContributed || same.remoteContributed).toBe(false)
	})

	it("keeps the newest gc time and merges unknown document fields", () => {
		const { doc } = mergeWatchState({
			local: docWith(
				{},
				{
					gc: 10,
					a: 1,
					reg: {
						t: 1,
						d: "x"
					}
				}
			),
			remote: docWith(
				{},
				{
					gc: 20,
					b: 2,
					reg: {
						t: 2,
						d: "x"
					}
				}
			)
		})

		expect(doc).toEqual({
			v: 1,
			gc: 20,
			items: {},
			series: {},
			a: 1,
			b: 2,
			reg: {
				t: 2,
				d: "x"
			}
		})
	})

	// Play count follows the register rule like every other field: two devices completing concurrently do not
	// add up, the later completion's count wins. It is a per-file statistic, so convergence beats exact sums.
	it("resolves concurrent completions to the winning register's play count", () => {
		const base = withCompletion({
			rec: undefined,
			stamp: stamp(10)
		})
		const onA = withCompletion({
			rec: base,
			stamp: stamp(20, "dev-a")
		})
		const onB = withCompletion({
			rec: base,
			stamp: stamp(21, "dev-b")
		})
		const { doc } = mergeWatchState({
			local: docWith({
				"s:a": onA
			}),
			remote: docWith({
				"s:a": onB
			})
		})

		expect(doc.items["s:a"]?.w).toEqual({
			t: 21,
			d: "dev-b",
			pl: true,
			pc: 2
		})
	})

	describe("properties over seeded random documents", () => {
		const random = prng(0x5eed)
		const pick = <T>(values: readonly T[]): T => values[Math.floor(random() * values.length)] as T
		const STAMPS = [1, 2, 3, 5, 8]
		const DEVICES = ["a", "b", "c"]
		const maybe = (probability: number): boolean => random() < probability
		const stampNow = (): Stamp => stamp(pick(STAMPS), pick(DEVICES))
		const unknownValue = (): unknown =>
			maybe(0.5)
				? {
						...stampNow(),
						x: pick([1, 2, "y"])
					}
				: pick([1, "two", [3], null, { four: 4 }])

		const randomRec = (): ItemRec => {
			let rec: ItemRec = {}

			if (maybe(0.6)) {
				rec = withProgress({
					rec,
					stamp: stampNow(),
					pos: pick([0, 100, 1800]),
					dur: 3600,
					lp: pick([1, 2, 3]),
					uuid: pick(["u1", "u2"]),
					size: 10
				})
			}

			if (maybe(0.5)) {
				rec = {
					...rec,
					w: {
						...stampNow(),
						pl: maybe(0.5),
						pc: pick([0, 1, 2])
					}
				}
			}

			if (maybe(0.3)) {
				rec = withSubtitleOffset({
					rec,
					stamp: stampNow(),
					subMs: pick([-100, 0, 250])
				})
			}

			if (maybe(0.3)) {
				rec = hideFromContinueWatching({
					rec,
					stamp: stampNow()
				})
			}

			if (maybe(0.2)) {
				rec[pick(["y", "z"])] = unknownValue()
			}

			if (maybe(0.2)) {
				rec = deleteRecord({
					rec,
					at: pick([1, 2])
				})

				if (maybe(0.5)) {
					rec = withSubtitleOffset({
						rec,
						stamp: stampNow(),
						subMs: 1
					})
				}
			}

			return rec
		}

		const randomDoc = (): WatchStateDoc => {
			const doc = emptyWatchState()

			for (const key of ["s:1", "s:2", "s:3", "s:4", "s:5"]) {
				const rec = randomRec()

				if (maybe(0.6) && Object.keys(rec).length > 0) {
					doc.items[key] = rec
				}
			}

			for (const key of ["d:1", "d:2"]) {
				let prefs: SeriesPrefs = {}

				if (maybe(0.5)) {
					prefs = withAudioLanguage({
						prefs,
						stamp: stampNow(),
						language: pick(["eng", "ger"])
					})
				}

				if (maybe(0.5)) {
					prefs = withSubtitleLanguage({
						prefs,
						stamp: stampNow(),
						language: pick([null, "eng"])
					})
				}

				if (maybe(0.2)) {
					prefs = deleteRecord({
						rec: prefs,
						at: 1
					})
				}

				if (Object.keys(prefs).length > 0) {
					doc.series[key] = prefs
				}
			}

			if (maybe(0.4)) {
				doc.gc = pick([1, 2, 3])
			}

			if (maybe(0.3)) {
				doc[pick(["k1", "k2"])] = unknownValue()
			}

			return doc
		}

		const merge = (local: WatchStateDoc, remote: WatchStateDoc): WatchStateDoc =>
			mergeWatchState({
				local,
				remote
			}).doc

		const RUNS = 500

		it("is commutative", () => {
			for (let i = 0; i < RUNS; i++) {
				const a = randomDoc()
				const b = randomDoc()

				expect(merge(a, b)).toEqual(merge(b, a))
			}
		})

		it("is associative", () => {
			for (let i = 0; i < RUNS; i++) {
				const a = randomDoc()
				const b = randomDoc()
				const c = randomDoc()

				expect(merge(merge(a, b), c)).toEqual(merge(a, merge(b, c)))
			}
		})

		it("is idempotent", () => {
			for (let i = 0; i < RUNS; i++) {
				const a = randomDoc()
				const b = randomDoc()
				const ab = merge(a, b)

				expect(merge(a, a)).toEqual(a)
				expect(merge(ab, b)).toEqual(ab)
				expect(merge(ab, a)).toEqual(ab)
			}
		})

		it("survives a JSON round trip through the parser", () => {
			for (let i = 0; i < RUNS; i++) {
				const a = merge(randomDoc(), randomDoc())

				expect(parsed(JSON.parse(JSON.stringify(a))).doc).toEqual(a)
			}
		})

		it("reports contributions consistently with the result", () => {
			for (let i = 0; i < RUNS; i++) {
				const local = randomDoc()
				const remote = randomDoc()
				const { doc, localContributed, remoteContributed } = mergeWatchState({
					local,
					remote
				})

				expect(localContributed).toBe(!isEqual(doc, remote))
				expect(remoteContributed).toBe(!isEqual(doc, local))
			}
		})
	})
})

function isEqual(a: unknown, b: unknown): boolean {
	try {
		expect(a).toEqual(b)

		return true
	} catch {
		return false
	}
}

describe("completion", () => {
	it("completes past 90% or within a second of the end", () => {
		expect(
			isCompletePosition({
				pos: 3240,
				dur: 3600,
				itemClass: "movie"
			})
		).toBe(false)
		expect(
			isCompletePosition({
				pos: 3241,
				dur: 3600,
				itemClass: "movie"
			})
		).toBe(true)
		expect(
			isCompletePosition({
				pos: 599,
				dur: 600,
				itemClass: "other"
			})
		).toBe(true)
	})

	it("completes short movies and episodes from 5%, but not other videos or extras", () => {
		for (const itemClass of ["movie", "episode"] as const) {
			expect(
				isCompletePosition({
					pos: 15,
					dur: 299,
					itemClass
				})
			).toBe(true)
			expect(
				isCompletePosition({
					pos: 14,
					dur: 299,
					itemClass
				})
			).toBe(false)
		}

		for (const itemClass of ["other", "extra"] as const) {
			expect(
				isCompletePosition({
					pos: 150,
					dur: 299,
					itemClass
				})
			).toBe(false)
			expect(
				isCompletePosition({
					pos: 270,
					dur: 299,
					itemClass
				})
			).toBe(true)
		}

		expect(
			isCompletePosition({
				pos: 15,
				dur: 300,
				itemClass: "movie"
			})
		).toBe(false)
	})

	it("completes an unknown duration only at the natural end", () => {
		expect(
			isCompletePosition({
				pos: 5000,
				dur: 0,
				itemClass: "movie"
			})
		).toBe(false)
		expect(
			isCompletePosition({
				pos: 5000,
				dur: 0,
				itemClass: "movie",
				ended: true
			})
		).toBe(true)
	})

	it("detects the crossing only once", () => {
		expect(
			crossesCompletion({
				from: 3200,
				to: 3300,
				dur: 3600,
				itemClass: "episode"
			})
		).toBe(true)
		expect(
			crossesCompletion({
				from: 3300,
				to: 3400,
				dur: 3600,
				itemClass: "episode"
			})
		).toBe(false)
	})
})

describe("deriveWatchState", () => {
	const file = {
		uuid: "file-1",
		size: 1000
	}

	const derive = (rec: ItemRec | undefined, over: Partial<typeof file & { durationSec: number }> = {}) =>
		deriveWatchState({
			rec,
			file: {
				...file,
				...over
			}
		})

	it("is resumable between 5% and 90% of a long enough item", () => {
		expect(isResumablePosition({ pos: 180, dur: 3600 })).toBe(true)
		expect(isResumablePosition({ pos: 179, dur: 3600 })).toBe(false)
		expect(isResumablePosition({ pos: 3240, dur: 3600 })).toBe(true)
		expect(isResumablePosition({ pos: 3241, dur: 3600 })).toBe(false)
		expect(isResumablePosition({ pos: 100, dur: 299 })).toBe(false)
		expect(isResumablePosition({ pos: 100, dur: 0 })).toBe(false)

		const state = derive(
			progress({
				t: 10,
				pos: 1800,
				lp: NOW - 5000
			})
		)

		expect(state).toEqual({
			played: false,
			playCount: 0,
			inProgress: true,
			hidden: false,
			resumable: true,
			positionSec: 1800,
			durationSec: 3600,
			lastPlayedMs: NOW - 5000
		})
	})

	it("is not resumable once a newer watched register exists, until played again", () => {
		const watched = markWatched({
			rec: progress({
				t: 10,
				pos: 1800
			}),
			stamp: stamp(11)
		})

		expect(derive(watched).resumable).toBe(false)
		expect(derive(watched).played).toBe(true)

		const rewatch = withProgress({
			rec: watched,
			stamp: stamp(12),
			pos: 900,
			dur: 3600,
			lp: NOW,
			uuid: "file-1",
			size: 1000
		})

		expect(derive(rewatch)).toMatchObject({
			played: true,
			resumable: true,
			positionSec: 900
		})
	})

	it("hides from Continue Watching until the next play", () => {
		const hidden = hideFromContinueWatching({
			rec: progress({
				t: 10,
				pos: 1800
			}),
			stamp: stamp(11)
		})

		expect(derive(hidden)).toMatchObject({
			inProgress: true,
			hidden: true,
			resumable: false,
			positionSec: 1800
		})

		const replayed = withProgress({
			rec: hidden,
			stamp: stamp(12),
			pos: 1900,
			dur: 3600,
			lp: NOW,
			uuid: "file-1",
			size: 1000
		})

		expect(derive(replayed)).toMatchObject({
			hidden: false,
			resumable: true
		})
	})

	it("checks the fingerprint of a replaced file", () => {
		const rec = progress({
			t: 10,
			pos: 1800,
			u: "old"
		})

		expect(derive(rec).resumable).toBe(true)
		expect(derive(rec, { size: 999 }).resumable).toBe(false)
		expect(derive(rec, { durationSec: 3602 }).resumable).toBe(true)
		expect(derive(rec, { durationSec: 3603 }).resumable).toBe(false)
		expect(
			derive(
				progress({
					t: 10,
					pos: 1800
				}),
				{
					size: 1,
					durationSec: 9000
				}
			).resumable
		).toBe(true)
	})

	it("prefers the probed duration", () => {
		expect(derive(progress({ t: 1, pos: 1800, dur: 0 })).resumable).toBe(false)
		expect(derive(progress({ t: 1, pos: 1800, dur: 0 }), { durationSec: 3600 }).resumable).toBe(true)
	})

	it("counts a mark-watched as a play time and caps fast clocks", () => {
		const rec = markWatched({
			rec: undefined,
			stamp: stamp((NOW + 1000) * 1000)
		})

		expect(derive(rec).lastPlayedMs).toBe(NOW + 1000)
		expect(
			deriveWatchState({
				rec,
				file,
				lastPlayedCeilingMs: NOW
			}).lastPlayedMs
		).toBe(NOW)
		expect(derive(undefined)).toEqual({
			played: false,
			playCount: 0,
			inProgress: false,
			hidden: false,
			resumable: false,
			positionSec: 0,
			durationSec: 0,
			lastPlayedMs: 0
		})
	})
})

describe("register writes", () => {
	it("records progress at millisecond precision", () => {
		const rec = withProgress({
			rec: undefined,
			stamp: stamp(1),
			pos: 12.3456789,
			dur: 100.0004,
			lp: 5,
			uuid: "u",
			size: 7,
			part: 2
		})

		expect(rec.p).toEqual({
			t: 1,
			d: "dev-a",
			pos: 12.346,
			dur: 100,
			lp: 5,
			u: "u",
			sz: 7,
			part: 2
		})
	})

	it("counts completions but not manual marks", () => {
		let rec = withCompletion({
			rec: undefined,
			stamp: stamp(1)
		})

		rec = withCompletion({
			rec,
			stamp: stamp(2)
		})
		expect(rec.w?.pc).toBe(2)

		rec = markWatched({
			rec,
			stamp: stamp(3)
		})
		expect(rec.w).toMatchObject({
			pl: true,
			pc: 2
		})
	})

	it("marks unwatched by resetting played, play count and position", () => {
		const rec = markUnwatched({
			rec: withCompletion({
				rec: progress({
					t: 1,
					pos: 3500
				}),
				stamp: stamp(2)
			}),
			stamp: stamp(3)
		})

		expect(rec.w).toEqual({
			t: 3,
			d: "dev-a",
			pl: false,
			pc: 0
		})
		expect(rec.p).toMatchObject({
			t: 3,
			pos: 0,
			lp: NOW
		})
		expect(
			markUnwatched({
				rec: undefined,
				stamp: stamp(1)
			}).p
		).toBeUndefined()
	})

	it("keeps audio and subtitle languages independent", () => {
		const onA = withAudioLanguage({
			prefs: undefined,
			stamp: stamp(1, "dev-a"),
			language: "jpn"
		})
		const onB = withSubtitleLanguage({
			prefs: undefined,
			stamp: stamp(2, "dev-b"),
			language: null
		})
		const merged = mergeWatchState({
			local: {
				...emptyWatchState(),
				series: {
					"d:s": onA
				}
			},
			remote: {
				...emptyWatchState(),
				series: {
					"d:s": onB
				}
			}
		}).doc.series["d:s"]

		expect(merged).toEqual({
			al: {
				t: 1,
				d: "dev-a",
				v: "jpn"
			},
			sl: {
				t: 2,
				d: "dev-b",
				v: null
			}
		})
		expect(
			withSubtitleLanguage({
				prefs: onA,
				stamp: stamp(3),
				language: "eng"
			}).al
		).toBe(onA.al)
		expect(seriesKey("x")).toBe("d:x")
		expect(itemKey("x")).toBe("s:x")
	})

	it("ignores invalid language codes and rejects invalid stamps or uuids", () => {
		expect(
			withAudioLanguage({
				prefs: undefined,
				stamp: stamp(1),
				language: ""
			})
		).toEqual({})
		expect(() =>
			markWatched({
				rec: undefined,
				stamp: stamp(Number.NaN)
			})
		).toThrow(RangeError)
		expect(() =>
			markWatched({
				rec: undefined,
				stamp: stamp(1, "")
			})
		).toThrow(RangeError)
		expect(() =>
			withProgress({
				rec: undefined,
				stamp: stamp(1),
				pos: 1,
				dur: 1,
				lp: 1,
				uuid: "",
				size: 1
			})
		).toThrow(RangeError)
	})

	it("produces only registers the parser accepts, whatever the inputs", () => {
		const random = prng(99)
		const HOSTILE = [
			Number.NaN,
			Infinity,
			-Infinity,
			-1,
			-0,
			-0.4,
			0,
			0.5,
			1.5,
			1e308,
			-1e308,
			2 ** 60,
			Number.MAX_SAFE_INTEGER,
			Number.MIN_VALUE,
			3600
		]
		const pick = <T>(values: readonly T[]): T => values[Math.floor(random() * values.length)] as T
		let t = 1

		const next = (): Stamp => stamp((t += 1 + Math.floor(random() * 1000)), pick(["dev-a", "dev-b"]))

		for (let i = 0; i < 1000; i++) {
			let rec: ItemRec = withProgress({
				rec: undefined,
				stamp: next(),
				pos: pick(HOSTILE),
				dur: pick(HOSTILE),
				lp: pick(HOSTILE),
				uuid: "u",
				size: pick(HOSTILE),
				part: pick(HOSTILE)
			})

			rec = withSubtitleOffset({
				rec,
				stamp: next(),
				subMs: pick(HOSTILE)
			})

			for (const write of [withCompletion, markWatched, markUnwatched, hideFromContinueWatching]) {
				if (random() < 0.5) {
					rec = write({
						rec,
						stamp: next()
					})
				}
			}

			if (random() < 0.2) {
				rec = {
					...rec,
					w: {
						...next(),
						pl: true,
						pc: Number.MAX_SAFE_INTEGER
					}
				}
				rec = withCompletion({
					rec,
					stamp: next()
				})
			}

			if (random() < 0.2) {
				rec = deleteRecord({
					rec,
					at: pick(HOSTILE)
				})
			}

			const prefs = withSubtitleLanguage({
				prefs: withAudioLanguage({
					prefs: undefined,
					stamp: next(),
					language: pick(["eng", "", "x".repeat(200)])
				}),
				stamp: next(),
				language: pick([null, "ger", ""])
			})
			const doc: WatchStateDoc = {
				...docWith({
					"s:a": rec
				}),
				series: Object.keys(prefs).length > 0 ? { "d:s": prefs } : {}
			}
			const result = parsed(JSON.parse(JSON.stringify(doc)))

			expect(result.dropped).toBe(0)
			expect(result.doc).toEqual(doc)
		}
	})
})

describe("writer ids", () => {
	it("never lose updates between distinct writers, even with equal clocks", () => {
		const t = NOW * 1000
		const tab1 = stamp(t, "device#tab-1")
		const tab2 = stamp(t, "device#tab-2")
		const one = docWith({
			"s:a": progress({
				t,
				pos: 600,
				d: tab1.d
			}),
			"s:b": withSubtitleOffset({
				rec: undefined,
				stamp: tab1,
				subMs: 100
			})
		})
		const two = docWith({
			"s:a": progress({
				t,
				pos: 900,
				d: tab2.d
			}),
			"s:b": hideFromContinueWatching({
				rec: undefined,
				stamp: tab2
			})
		})
		const ab = mergeWatchState({
			local: one,
			remote: two
		}).doc
		const ba = mergeWatchState({
			local: two,
			remote: one
		}).doc

		expect(ab).toEqual(ba)
		// Different registers both survive; on the same register the stamp order decides, visibly to both writers
		expect(Object.keys(ab.items["s:b"] ?? {}).sort()).toEqual(["h", "o"])
		expect(ab.items["s:a"]?.p?.d).toBe(tab2.d)

		const reasserted = progress({
			t: hlcTick({
				clock: hlcObserve({
					clock: t,
					seen: t,
					now: NOW
				}),
				now: NOW
			}),
			pos: 600,
			d: tab1.d
		})
		const after = mergeWatchState({
			local: docWith({
				"s:a": reasserted
			}),
			remote: ab
		}).doc

		expect(after.items["s:a"]?.p?.pos).toBe(600)
		expect(compareStamps(tab1, tab2)).not.toBe(0)
	})
})

describe("compactWatchState", () => {
	const sizeOf = (doc: WatchStateDoc): number => utf8Length(JSON.stringify(doc))
	const at = (ms: number): number => ms * 1000

	it("returns the document itself when it fits", () => {
		const doc = docWith({
			"s:a": progress({
				t: 1,
				pos: 10
			})
		})
		const result = compactWatchState({
			doc
		})

		expect(result.doc).toBe(doc)
		expect(result.bytes).toBe(sizeOf(doc))
		expect(result.removed).toBe(0)
	})

	it("tombstones stale unplayed progress, then superseded progress, then whole records, against the document's own clock", () => {
		// The document's newest stamp is a day before NOW; wall clocks play no part.
		const ref = NOW - DAY
		const items: Record<string, ItemRec> = {
			"s:stale": progress({
				t: at(ref - STALE_PROGRESS_MS - DAY),
				pos: 100,
				lp: ref - STALE_PROGRESS_MS - DAY
			}),
			"s:superseded": withCompletion({
				rec: progress({
					t: at(ref - 2 * DAY),
					pos: 3500,
					lp: ref - 2 * DAY
				}),
				stamp: stamp(at(ref - 2 * DAY) + 1)
			}),
			"s:recent": progress({
				t: at(ref),
				pos: 100,
				lp: ref
			}),
			"s:old": markWatched({
				rec: undefined,
				stamp: stamp(at(ref - 100 * DAY))
			})
		}
		const doc = docWith(items)
		const first = compactWatchState({
			doc,
			limitBytes: sizeOf(doc) - 1
		})

		expect(first.bytes).toBe(sizeOf(first.doc))
		expect(first.doc.items["s:stale"]).toEqual({
			x: {
				t: at(ref - STALE_PROGRESS_MS - DAY),
				d: "dev-a",
				e: ref
			}
		})
		expect(first.doc.items["s:superseded"]).toBe(items["s:superseded"])

		const second = compactWatchState({
			doc,
			limitBytes: first.bytes - 1
		})

		expect(second.doc.items["s:superseded"]?.p).toBeUndefined()
		expect(second.doc.items["s:superseded"]?.w).toEqual(items["s:superseded"]?.w)
		expect(second.doc.items["s:recent"]).toBe(items["s:recent"])

		// Out of reach of the soft target: whole records only go to respect the hard limit
		expect(
			compactWatchState({
				doc,
				limitBytes: 1
			}).doc.items["s:recent"]
		).toBe(items["s:recent"])

		const hardOnly = compactWatchState({
			doc,
			limitBytes: 1,
			hardLimitBytes: sizeOf(second.doc) - 1
		})

		// Oldest first: the superseded record goes whole, the recent one stays
		expect(Object.keys(hardOnly.doc.items["s:superseded"] ?? {})).toEqual(["x"])
		expect(hardOnly.doc.items["s:recent"]).toBe(items["s:recent"])

		const tight = compactWatchState({
			doc,
			limitBytes: 1,
			hardLimitBytes: 1
		})

		// Tombstone-only records are dropped last; a tombstone would outweigh the lone watched register
		expect(tight.doc.items).toEqual({
			"s:old": items["s:old"]
		})
		expect(doc.items).toEqual(items)
		expect(
			compactWatchState({
				doc: tight.doc,
				limitBytes: 1,
				hardLimitBytes: 1
			}).doc
		).toBe(tight.doc)
	})

	it("expires tombstone-only records after the TTL measured on the document's own clock", () => {
		const doc = docWith({
			"s:live": progress({
				t: at(NOW),
				pos: 100
			}),
			"s:expired": {
				x: {
					t: 1,
					d: "dev-a",
					e: NOW - TOMBSTONE_TTL_MS - 1
				}
			},
			"s:kept": {
				x: {
					t: 1,
					d: "dev-a",
					e: NOW - TOMBSTONE_TTL_MS + 1
				}
			}
		})

		expect(
			Object.keys(
				compactWatchState({
					doc
				}).doc.items
			)
		).toEqual(["s:live", "s:kept"])
	})

	it("prunes registers covered by a tombstone in every merge order and keeps later writes", () => {
		const original = progress({
			t: 5,
			pos: 100
		})
		const deleted = docWith({
			"s:a": deleteRecord({
				rec: original,
				at: NOW
			})
		})
		const stale = docWith({
			"s:a": original
		})
		const later = docWith({
			"s:a": withSubtitleOffset({
				rec: undefined,
				stamp: stamp(6),
				subMs: 1
			})
		})

		for (const [a, b] of [
			[deleted, stale],
			[stale, deleted]
		] as const) {
			expect(
				mergeWatchState({
					local: a,
					remote: b
				}).doc.items["s:a"]
			).toEqual(deleted.items["s:a"])
		}

		expect(
			Object.keys(
				mergeWatchState({
					local: later,
					remote: deleted
				}).doc.items["s:a"] ?? {}
			).sort()
		).toEqual(["o", "x"])
		expect(
			mergeWatchState({
				local: stale,
				remote: deleted
			}).localContributed
		).toBe(false)
	})

	it("never stops above the limit it can reach and reports the exact size", () => {
		const items: Record<string, ItemRec> = {}

		for (let i = 0; i < 40; i++) {
			items[itemKey(`item-${i}`)] = progress({
				t: at(NOW - i * DAY),
				pos: 600,
				lp: NOW - i * DAY,
				u: `file-${i}`
			})
		}

		const doc = docWith(items)
		const empty = sizeOf(emptyWatchState())

		for (let limit = empty; limit <= sizeOf(doc); limit += 37) {
			const result = compactWatchState({
				doc,
				limitBytes: limit,
				hardLimitBytes: limit
			})

			expect(result.bytes).toBe(sizeOf(result.doc))
			expect(result.bytes).toBeLessThanOrEqual(limit)
		}
	})

	it("converges across replicas flushing in any order and settles without oscillating", () => {
		const random = prng(42)
		const limits = {
			limitBytes: 1500,
			hardLimitBytes: 3000
		}

		for (let run = 0; run < 40; run++) {
			let head = emptyWatchState()
			const replicas = ["r1#t", "r2#t", "r3#t"].map(d => ({
				d,
				clock: 0,
				now: NOW,
				doc: emptyWatchState()
			}))

			const flush = (replica: (typeof replicas)[number]): boolean => {
				const result = reconcileWatchState({
					local: replica.doc,
					remote: head,
					...limits
				})
				const swapped = reconcileWatchState({
					local: head,
					remote: replica.doc,
					...limits
				})

				expect(result.doc).toEqual(swapped.doc)
				expect(result.bytes).toBeLessThanOrEqual(limits.hardLimitBytes)

				const reread = parsed(JSON.parse(JSON.stringify(result.doc)))

				expect(reread.doc).toEqual(result.doc)
				replica.clock = hlcObserve({
					clock: replica.clock,
					seen: reread.maxStamp,
					now: replica.now
				})

				if (result.upload) {
					head = result.doc
				}

				replica.doc = result.doc

				return result.upload
			}

			for (let step = 0; step < 60; step++) {
				const replica = replicas[Math.floor(random() * replicas.length)] as (typeof replicas)[number]

				if (random() < 0.3) {
					flush(replica)

					continue
				}

				replica.now += Math.floor(random() * 3 * DAY)
				replica.clock = hlcTick({
					clock: replica.clock,
					now: replica.now
				})

				const key = itemKey(`v${Math.floor(random() * 30)}`)
				const rec = replica.doc.items[key]
				const writeStamp = stamp(replica.clock, replica.d)
				const choice = random()
				const next =
					choice < 0.5
						? withProgress({
								rec,
								stamp: writeStamp,
								pos: 600,
								dur: 3600,
								lp: replica.now,
								uuid: `file-${key}`,
								size: 1
							})
						: choice < 0.85 || rec === undefined
							? withCompletion({
									rec,
									stamp: writeStamp
								})
							: deleteRecord({
									rec,
									at: replica.now
								})

				replica.doc = {
					...replica.doc,
					items: {
						...replica.doc.items,
						[key]: next
					}
				}
			}

			let rounds = 0

			while (replicas.map(flush).some(Boolean)) {
				rounds++
				expect(rounds).toBeLessThan(8)
			}

			for (const replica of replicas) {
				expect(replica.doc).toEqual(head)
			}
		}
	})

	it("uploads only when the head would change", () => {
		const remote = docWith({
			"s:a": progress({
				t: 5,
				pos: 100
			})
		})
		const stale = docWith({
			"s:a": progress({
				t: 4,
				pos: 50
			})
		})

		expect(
			reconcileWatchState({
				local: stale,
				remote
			})
		).toMatchObject({
			doc: remote,
			upload: false
		})
		expect(
			reconcileWatchState({
				local: structuredClone(remote),
				remote
			}).upload
		).toBe(false)
		expect(
			reconcileWatchState({
				local: remote,
				remote: stale
			}).upload
		).toBe(true)
	})

	it("keeps a document over the hard limit within it, newest records intact", () => {
		const items: Record<string, ItemRec> = {}

		for (let i = 0; i < 8000; i++) {
			const lp = NOW - i * 60_000
			const rec = progress({
				t: lp * 1000,
				pos: 1000 + i,
				lp,
				u: `00000000-0000-4000-8000-${String(i).padStart(12, "0")}`
			})

			items[itemKey(`10000000-0000-4000-8000-${String(i).padStart(12, "0")}`)] =
				i % 2 === 0
					? withCompletion({
							rec,
							stamp: stamp(lp * 1000 + 1)
						})
					: rec
		}

		const doc = docWith(items)
		const started = performance.now()
		const result = compactWatchState({
			doc
		})
		const elapsed = performance.now() - started

		console.info(`compactWatchState: ${sizeOf(doc)} B -> ${result.bytes} B, ${result.removed} tombstones in ${elapsed.toFixed(1)} ms`)
		expect(sizeOf(doc)).toBeGreaterThan(WATCH_STATE_HARD_LIMIT_BYTES)
		expect(result.bytes).toBeLessThanOrEqual(WATCH_STATE_HARD_LIMIT_BYTES)
		expect(result.bytes).toBe(sizeOf(result.doc))
		expect(result.doc.items[itemKey("10000000-0000-4000-8000-000000000001")]).toEqual(
			items[itemKey("10000000-0000-4000-8000-000000000001")]
		)
		expect(elapsed).toBeLessThan(2000)
	})

	it("merges two ~1 MiB documents quickly", () => {
		const build = (device: string, offset: number): WatchStateDoc => {
			const items: Record<string, ItemRec> = {}

			for (let i = 0; i < 6000; i++) {
				items[itemKey(`10000000-0000-4000-8000-${String(i).padStart(12, "0")}`)] = progress({
					t: NOW * 1000 + ((i + offset) % 7),
					pos: 600 + i,
					lp: NOW,
					d: device,
					u: `00000000-0000-4000-8000-${String(i).padStart(12, "0")}`
				})
			}

			return docWith(items)
		}

		const local = JSON.stringify(build("dev-a", 0))
		const remote = JSON.stringify(build("dev-b", 3))
		const started = performance.now()
		const a = parsed(JSON.parse(local)).doc
		const b = parsed(JSON.parse(remote)).doc
		const merged = mergeWatchState({
			local: a,
			remote: b
		})
		const elapsed = performance.now() - started

		console.info(`parse + merge of 2 x ${(local.length / 1024).toFixed(0)} KiB: ${elapsed.toFixed(1)} ms`)
		expect(merged.localContributed && merged.remoteContributed).toBe(true)
		expect(elapsed).toBeLessThan(1000)
	})
})
