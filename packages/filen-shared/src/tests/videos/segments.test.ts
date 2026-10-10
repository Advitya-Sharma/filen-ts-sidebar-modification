import { describe, expect, it } from "vitest"
import {
	chapterNameKind,
	defaultSkipMode,
	mergeSegments,
	parseEdl,
	segmentAction,
	segmentAt,
	segmentKindFromSkipType,
	segmentsFromChapters,
	skipTarget,
	type SkipSegment
} from "@filen/shared/videos"

describe("segmentKindFromSkipType", () => {
	it.each([
		[0, null],
		[1, "intro"],
		[2, "outro"],
		[3, "recap"],
		[4, "preview"],
		[5, "preview"],
		[6, "commercial"],
		[7, "intermission"],
		[8, null],
		[-1, null]
	])("%s -> %s", (value, kind) => {
		expect(segmentKindFromSkipType(value)).toBe(kind)
	})
})

describe("chapterNameKind", () => {
	it.each([
		["Intro", "intro"],
		["Opening", "intro"],
		["Opening Credits", "intro"],
		["OP", "intro"],
		["OP2", "intro"],
		["Vorspann", "intro"],
		["Générique de début", "intro"],
		["オープニング", "intro"],
		["Outro", "outro"],
		["Ending", "outro"],
		["End Credits", "outro"],
		["Credits", "outro"],
		["ED", "outro"],
		["Abspann", "outro"],
		["Générique de fin", "outro"],
		["Créditos", "outro"],
		["エンディング", "outro"],
		["Recap", "recap"],
		["Previously on", "recap"],
		["Précédemment", "recap"],
		["Preview", "preview"],
		["Next Time", "preview"],
		["Vorschau", "preview"],
		["Opening Theme", "intro"],
		["オープニング(テーマ)", "intro"],
		["オープニングテーマ", "intro"],
		["片头曲", "intro"],
		["片頭", "intro"],
		["오프닝", "intro"],
		["1. Intro", "intro"],
		["Sigla finale", "outro"],
		["Titoli di coda", "outro"],
		["片尾曲", "outro"],
		["エンディング(テーマ)", "outro"],
		["엔딩", "outro"],
		["Créditos finais", "outro"],
		["Ending Theme", "outro"],
		["Abertura", "intro"],
		["Apertura", "intro"],
		["Anteprima", "preview"],
		["次回予告", "preview"],
		["前情提要", "recap"],
		["Previously on Lost", "recap"]
	])("%s -> %s", (title, kind) => {
		expect(chapterNameKind(title)).toBe(kind)
	})

	it.each([
		"Chapter 1",
		"Interlude",
		"Introspection",
		"Operation",
		"Opened",
		"Previewing",
		"Scene 03",
		"Credit card",
		"",
		"   ",
		"Post-Credits Scene",
		"Mid-Credits Scene",
		"Mid-credits",
		"Opening Scene",
		"The Ending",
		"The Opening",
		"Teaser",
		"Final Credits Discussion",
		"Der Abspann des Lebens",
		"Le générique est long",
		"Resumen del caso",
		"La sigla del gruppo",
		"片头的故事",
		"エンディングを迎える",
		"오프닝 장면"
	])("does not match %j (whole-title forms only)", title => {
		expect(chapterNameKind(title)).toBeNull()
	})

	it("handles null and long names without blowing up", () => {
		expect(chapterNameKind(null)).toBeNull()
		expect(chapterNameKind(undefined)).toBeNull()
		expect(chapterNameKind(`${"a ".repeat(100_000)}intro`)).toBeNull()
		expect(chapterNameKind(" ".repeat(100_000))).toBeNull()
	})
})

describe("segmentsFromChapters", () => {
	it("derives segments from skip types first, then names; chapters end at the next start", () => {
		const out = segmentsFromChapters(
			[
				{ start: 0, title: "Recap" },
				{ start: 30, skipType: 1, title: "something else" },
				{ start: 120, title: "Act 1" },
				{ start: 1300, skipType: 2 },
				{ start: 1380, end: 1400, title: "Credits" },
				{ start: 1400, title: "Preview" }
			],
			1450
		)

		expect(out).toEqual([
			{ kind: "recap", start: 0, end: 30, source: "chapter" },
			{ kind: "intro", start: 30, end: 120, source: "skipType" },
			{ kind: "outro", start: 1300, end: 1380, source: "skipType" },
			{ kind: "outro", start: 1380, end: 1400, source: "chapter" },
			{ kind: "preview", start: 1400, end: 1450, source: "chapter" }
		])
	})

	it("honours enabled/hidden flags and ignores unknown skip types and zero-length chapters", () => {
		expect(
			segmentsFromChapters(
				[
					{ start: 0, skipType: 1, enabled: false },
					{ start: 10, skipType: 1, hidden: true },
					{ start: 20, skipType: 99, title: "Intro" },
					{ start: 30, skipType: 1 },
					{ start: 30, skipType: 2 }
				],
				100
			)
		).toEqual([{ kind: "outro", start: 30, end: 100, source: "skipType" }])
	})

	it("treats an explicit skip type 0 as never skip, without falling back to the title", () => {
		expect(segmentsFromChapters([{ start: 0, end: 60, skipType: 0, title: "Intro" }], 100)).toEqual([])
		expect(segmentsFromChapters([{ start: 0, end: 60, skipType: null, title: "Intro" }], 100)).toHaveLength(1)
		expect(segmentsFromChapters([{ start: 0, end: 60, title: "Intro" }], 100)).toHaveLength(1)
	})

	it("clamps to the duration", () => {
		expect(segmentsFromChapters([{ start: 90, skipType: 2, end: 500 }], 100)).toEqual([
			{ kind: "outro", start: 90, end: 100, source: "skipType" }
		])
	})
})

describe("parseEdl", () => {
	it("accepts MM:SS and HH:MM:SS(.sss) times", () => {
		expect(parseEdl("01:30 02:00.5 3\n1:00:05 1:00:10.250 3\n")).toEqual([
			{ kind: "commercial", start: 90, end: 120.5, source: "edl" },
			{ kind: "commercial", start: 3605, end: 3610.25, source: "edl" }
		])
		expect(parseEdl("1:2:3:4 9:9:9:9 3\n1.5:00 2:00 3\nx:10 20 3\n")).toEqual([])
	})

	it("parses 3-field lines (seconds), tabs and mixed spacing; cuts are mandatory, commercial breaks are not", () => {
		expect(parseEdl("10.5 20.25 0\n30\t40 3\n", null)).toEqual([
			{ kind: "commercial", start: 10.5, end: 20.25, source: "edl", mustSkip: true },
			{ kind: "commercial", start: 30, end: 40, source: "edl" }
		])
	})

	it("ignores mute and scene-marker actions", () => {
		expect(parseEdl("1 2 1\n3 4 2\n5 6 9\n")).toEqual([])
	})

	it("treats a 2-field line as time + action (zero length)", () => {
		expect(parseEdl("123.4 0\n200 3\n")).toEqual([])
	})

	it("needs fps for #frame tokens", () => {
		expect(parseEdl("#250 #500 0\n")).toEqual([])
		expect(parseEdl("#250 #500 0\n", 25)).toEqual([{ kind: "commercial", start: 10, end: 20, source: "edl", mustSkip: true }])
		expect(parseEdl("#250 30 0\n", 25)).toEqual([{ kind: "commercial", start: 10, end: 30, source: "edl", mustSkip: true }])
		expect(parseEdl("#250 #500 0\n", 0)).toEqual([])
		expect(parseEdl("10 #500 0\n")).toEqual([])
	})

	it("drops malformed, negative, inverted and comment lines", () => {
		expect(parseEdl("\n// note\n# note\nabc def 0\n-1 5 0\n20 10 0\n1 2\n1 2 3 4\n1\n   \r\n")).toEqual([])
		expect(parseEdl("5 6 0\r\n")).toHaveLength(1)
	})

	it("caps the number of lines", () => {
		const text = "1 2 0\n".repeat(50_000)

		expect(parseEdl(text).length).toBe(10_000)
	})
})

describe("mergeSegments", () => {
	const seg = (kind: SkipSegment["kind"], start: number, end: number, source: SkipSegment["source"]): SkipSegment => ({
		kind,
		start,
		end,
		source
	})

	it("lets higher-priority sources shadow overlapping lower ones and sorts the result", () => {
		const merged = mergeSegments(
			[
				[seg("commercial", 5, 15, "edl"), seg("commercial", 500, 520, "edl")],
				[seg("intro", 10, 40, "chapter")],
				[seg("intro", 0, 12, "skipType")]
			],
			1000
		)

		expect(merged).toEqual([seg("intro", 0, 12, "skipType"), seg("commercial", 500, 520, "edl")])
	})

	it("keeps must-skip cuts whole, joins overlapping ones and trims every other segment around them", () => {
		const cut = (start: number, end: number): SkipSegment => ({ ...seg("commercial", start, end, "edl"), mustSkip: true })
		const merged = mergeSegments(
			[
				[seg("intro", 590, 605, "chapter"), seg("outro", 900, 1000, "chapter"), seg("recap", 100, 110, "chapter")],
				[cut(600, 660), cut(650, 700), cut(950, 960), cut(95, 120)],
				[seg("preview", 300, 400, "skipType")]
			],
			1000
		)

		expect(merged).toEqual([
			cut(95, 120),
			seg("preview", 300, 400, "skipType"),
			seg("intro", 590, 600, "chapter"),
			cut(600, 700),
			seg("outro", 900, 950, "chapter"),
			cut(950, 960),
			seg("outro", 960, 1000, "chapter")
		])
		expect(segmentAction(merged[3] as SkipSegment, "off", 1000)).toBe("autoskip")
		expect(segmentAction(merged[6] as SkipSegment, "ask", 1000)).toBe("upNext")
	})

	it("drops a lower-priority segment overlapping a kept one whether or not a cut splits it", () => {
		const cut: SkipSegment = { ...seg("commercial", 100, 110, "edl"), mustSkip: true }
		const intro = seg("intro", 0, 90, "skipType")
		const sources = [[intro], [seg("recap", 50, 150, "chapter"), seg("outro", 85, 200, "chapter")]]

		expect(mergeSegments(sources, 1000)).toEqual([intro])
		expect(mergeSegments([...sources, [cut]], 1000)).toEqual([intro, cut])
	})

	it("merges thousands of cuts and segments spanning them in near-linear time", () => {
		const cuts: SkipSegment[] = Array.from({ length: 3000 }, (_, index) => ({
			...seg("commercial", index * 10 + 1, index * 10 + 2, "edl"),
			mustSkip: true
		}))
		const spanning = Array.from({ length: 7000 }, () => seg("commercial", 0, 30_000, "edl"))
		const started = performance.now()
		const merged = mergeSegments([cuts, spanning], 30_000)

		expect(performance.now() - started).toBeLessThan(1000)
		expect(merged.length).toBe(3000 + 3001)
		expect(merged[0]).toEqual(seg("commercial", 0, 1, "edl"))
		expect(merged[1]).toEqual(cuts[0])
		expect(merged.at(-1)).toEqual(seg("commercial", 29_992, 30_000, "edl"))
	})

	it("clamps and drops degenerate segments", () => {
		expect(
			mergeSegments([[seg("outro", 90, 200, "chapter"), seg("intro", 5, 5, "chapter"), seg("intro", -5, 3, "chapter")]], 100)
		).toEqual([seg("intro", 0, 3, "chapter"), seg("outro", 90, 100, "chapter")])
	})

	it("is empty for no input", () => {
		expect(mergeSegments([], 100)).toEqual([])
	})
})

describe("segmentAt", () => {
	const list: SkipSegment[] = [
		{ kind: "intro", start: 10, end: 20, source: "chapter" },
		{ kind: "outro", start: 100, end: 120, source: "chapter" }
	]

	it("finds the containing segment (end exclusive)", () => {
		expect(segmentAt(list, 10)?.kind).toBe("intro")
		expect(segmentAt(list, 19.99)?.kind).toBe("intro")
		expect(segmentAt(list, 20)).toBeNull()
		expect(segmentAt(list, 110)?.kind).toBe("outro")
		expect(segmentAt(list, 5)).toBeNull()
		expect(segmentAt([], 5)).toBeNull()
	})
})

describe("prompt rules", () => {
	const seg = (kind: SkipSegment["kind"], start: number, end: number): SkipSegment => ({ kind, start, end, source: "chapter" })

	it("always skips EDL cuts, whatever the mode", () => {
		const cut: SkipSegment = { kind: "commercial", start: 100, end: 160, source: "edl", mustSkip: true }

		expect(segmentAction(cut, "off", 1000)).toBe("autoskip")
		expect(segmentAction(cut, "ask", 1000)).toBe("autoskip")
		expect(segmentAction({ ...cut, end: 100.5 }, "ask", 1000)).toBe("none")
		expect(segmentAction({ ...cut, mustSkip: false }, "off", 1000)).toBe("none")
	})

	it("defaults to ask for intro, outro and advertisements, off otherwise", () => {
		expect(defaultSkipMode("intro")).toBe("ask")
		expect(defaultSkipMode("outro")).toBe("ask")
		expect(defaultSkipMode("commercial")).toBe("ask")
		for (const kind of ["recap", "preview", "intermission"] as const) {
			expect(defaultSkipMode(kind)).toBe("off")
		}
	})

	it.each([
		["off", seg("intro", 0, 60), "none"],
		["ask", seg("intro", 0, 60), "prompt"],
		["ask", seg("intro", 0, 2.9), "none"],
		["ask", seg("intro", 0, 3), "prompt"],
		["auto", seg("intro", 0, 0.9), "none"],
		["auto", seg("intro", 0, 1), "autoskip"],
		["auto", seg("intro", 0, 2), "autoskip"],
		["ask", seg("outro", 1300, 1400), "prompt"],
		["ask", seg("outro", 1300, 1440), "upNext"],
		["ask", seg("outro", 1300, 1439.5), "upNext"],
		["auto", seg("outro", 1300, 1440), "upNext"],
		["off", seg("outro", 1300, 1440), "none"],
		["ask", seg("intro", 1300, 1440), "prompt"]
	] as const)("%s %j -> %s", (mode, segment, expected) => {
		expect(segmentAction(segment, mode, 1440)).toBe(expected)
	})

	it("seeks on click only when more than 1 s remains", () => {
		const segment = seg("intro", 10, 70)

		expect(skipTarget(segment, 12)).toBe(70)
		expect(skipTarget(segment, 68.9)).toBe(70)
		expect(skipTarget(segment, 69)).toBeNull()
		expect(skipTarget(segment, 69.5)).toBeNull()
	})
})
