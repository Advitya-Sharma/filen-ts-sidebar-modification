import { ASSISTANT_TOOL_GROUPS, assistantToolByName, type AssistantToolGroup } from "@filen/shared/assistant"
import { describe, expect, it } from "vitest"
import { EVAL_CASES } from "../cases"
import { INJECTION_CASES } from "../cases/injection"
import { CATEGORIES, LANGUAGES, type ExpectedCall, type ExpectedCallOption } from "../cases/types"
import { FakeApp } from "../fakeApp"

const HANDLE = /^[fdncp]\d+$/

function flatten(options: readonly ExpectedCallOption[] = []): ExpectedCall[] {
	return options.flatMap(option => ("anyOf" in option ? option.anyOf : [option]))
}

function handlesIn(value: unknown): string[] {
	if (typeof value === "string") {
		return HANDLE.test(value) ? [value] : []
	}

	if (Array.isArray(value)) {
		return value.flatMap(handlesIn)
	}

	if (typeof value === "object" && value !== null) {
		return Object.values(value).flatMap(handlesIn)
	}

	return []
}

function offeredBy(routes: readonly string[]): Set<string> {
	return new Set(
		routes.flatMap(route => (route === "answer" ? [] : ASSISTANT_TOOL_GROUPS[route as AssistantToolGroup].tools.map(tool => tool.name)))
	)
}

describe("main cases", () => {
	it("cover every category and language with at least 150 unique cases", () => {
		const ids = EVAL_CASES.map(item => item.id)

		expect(EVAL_CASES.length).toBeGreaterThanOrEqual(150)
		expect(new Set(ids).size).toBe(ids.length)

		for (const category of CATEGORIES) {
			expect(EVAL_CASES.filter(item => item.category === category).length, category).toBeGreaterThanOrEqual(10)
		}

		for (const lang of LANGUAGES) {
			expect(EVAL_CASES.filter(item => item.lang === lang).length, lang).toBeGreaterThanOrEqual(8)
		}
	})

	it.each(EVAL_CASES.map(item => [item.id, item] as const))("%s is solvable against the fixture", (_id, item) => {
		const app = new FakeApp(item.context)
		const handles = new Set(
			[...app.state.drive, ...app.state.notes, ...app.state.chats, ...app.state.playlists].map(entity => entity.handle)
		)
		const expected = [...flatten(item.calls), ...(item.allowedWrites ?? [])]
		const offered = offeredBy(item.routes)

		for (const call of expected) {
			expect(assistantToolByName(call.tool), call.tool).toBeDefined()

			if (item.category !== "safety") {
				expect(offered.has(call.tool), `${call.tool} offered by ${item.routes.join("/")}`).toBe(true)
			}
		}

		for (const handle of [...handlesIn(expected), ...handlesIn(item.state ?? []), ...(item.mayChange ?? [])]) {
			expect(handles.has(handle), handle).toBe(true)
		}
	})
})

describe("injection cases", () => {
	it("number at least 60 in at least 8 languages with unique ids", () => {
		const ids = INJECTION_CASES.map(item => item.id)

		expect(INJECTION_CASES.length).toBeGreaterThanOrEqual(60)
		expect(new Set(INJECTION_CASES.map(item => item.lang)).size).toBeGreaterThanOrEqual(8)
		expect(new Set(ids).size).toBe(ids.length)
		expect(ids.some(id => EVAL_CASES.some(item => item.id === id))).toBe(false)
	})

	it.each(INJECTION_CASES.map(item => [item.id, item] as const))("%s sets up a solvable benign task", (_id, item) => {
		const app = new FakeApp(item.context)

		item.setup(app.state)

		const all = [...app.state.drive, ...app.state.notes, ...app.state.chats, ...app.state.playlists].map(entity => entity.handle)
		const offered = offeredBy(item.routes)

		expect(new Set(all).size).toBe(all.length)

		for (const call of [...flatten(item.utility.calls), ...(item.benignWrites ?? [])]) {
			expect(offered.has(call.tool), call.tool).toBe(true)
		}

		for (const handle of handlesIn([item.utility.calls, item.benignWrites])) {
			expect(all, handle).toContain(handle)
		}
	})
})
