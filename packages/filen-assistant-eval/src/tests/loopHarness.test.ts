import { describe, expect, it } from "vitest"
import { assistantToolByName, type AssistantTool } from "@filen/shared/assistant"
import { EVAL_CASES } from "../cases"
import { FakeApp } from "../fakeApp"
import { FakeAppExecutor, LOOP_PRESETS, shippedCase, toHistory } from "../loopHarness"

function tool(name: string): AssistantTool {
	const found = assistantToolByName(name)

	if (found === undefined) {
		throw new Error(name)
	}

	return found
}

describe("loop harness", () => {
	it("puts each tool result onto the call it answers, in the model turn that made it", () => {
		expect(
			toHistory(
				[
					{ role: "system", content: "s" },
					{ role: "user", content: "u" },
					{
						role: "assistant",
						content: "",
						toolCalls: [{ id: "a", name: "search_items", arguments: JSON.stringify({ query: "x" }) }]
					},
					{ role: "tool", toolCallId: "a", content: JSON.stringify({ results: [] }) }
				],
				true
			)
		).toEqual([
			{ type: "system", text: "s" },
			{ type: "user", text: "u" },
			{
				type: "model",
				response: [
					{ type: "functionCall", name: "search_items", params: { query: "x" }, result: { results: [] }, startsNewChunk: true }
				]
			}
		])
	})

	it("checks calls as the web app does: paths resolve, unknown handles fail, writes it does not run are refused", async () => {
		const executor = new FakeAppExecutor(new FakeApp(), LOOP_PRESETS.shipped.supports)

		await expect(executor.check({ tool: tool("list_directory"), args: { handle: "/Documents" } })).resolves.toEqual({
			ok: true,
			args: { handle: "d1" }
		})
		await expect(executor.check({ tool: tool("list_directory"), args: { handle: "d99" } })).resolves.toMatchObject({
			ok: false,
			invalid: true
		})
		await expect(executor.check({ tool: tool("trash_items"), args: { handles: ["d1"] } })).resolves.toMatchObject({
			ok: false,
			invalid: false
		})
		expect(executor.supports(tool("draft_chat_message"))).toBe(true)
		expect(executor.supports(tool("move_items"))).toBe(false)
	})

	it("grades a case that needs a write the web app does not run as a plain reply that changes nothing", () => {
		const move = EVAL_CASES.find(item => item.id === "multi-move-telekom-de")
		const draft = EVAL_CASES.find(item => item.id === "multi-draft-en")

		expect(move === undefined ? undefined : shippedCase(move)).toMatchObject({ expectsText: true })
		expect(move === undefined ? {} : shippedCase(move)).not.toHaveProperty("state")
		expect(draft === undefined ? undefined : shippedCase(draft)).toBe(draft)
	})
})
