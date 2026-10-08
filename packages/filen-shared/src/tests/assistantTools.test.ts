import { describe, expect, it } from "vitest"
import * as z from "zod"
import {
	ASSISTANT_ROUTES,
	ASSISTANT_TOOL_GROUPS,
	ASSISTANT_TOOLS,
	assistantRouteDescription,
	defineTool,
	portableSchemaViolations,
	readChat,
	isConsequentialTool,
	toolInputJsonSchema
} from "@filen/shared/assistant"

describe("assistant tool schemas", () => {
	it.each(ASSISTANT_TOOLS.map(tool => [tool.name, tool] as const))("%s compiles to the portable subset", (_name, tool) => {
		const schema = toolInputJsonSchema(tool)

		expect(portableSchemaViolations(schema)).toEqual([])

		for (const [property, definition] of Object.entries(schema.properties)) {
			expect(
				definition.description ?? (definition.type === "string" && "enum" in definition),
				`${tool.name}.${property}`
			).toBeTruthy()
		}
	})

	it("drops the implicit safe-integer bounds of z.int()", () => {
		expect(toolInputJsonSchema(readChat).properties["limit"]).toEqual({
			type: "integer",
			description: "How many recent messages, default 20"
		})
	})

	it.each([
		[
			"a nested object",
			z.strictObject({
				nested: z.strictObject({
					a: z.string()
				})
			})
		],
		[
			"a union",
			z.strictObject({
				a: z.union([z.string(), z.boolean()]).describe("a")
			})
		],
		[
			"a number bound",
			z.strictObject({
				a: z.int().min(1).describe("a")
			})
		],
		[
			"a string format",
			z.strictObject({
				a: z.email().describe("a")
			})
		],
		[
			"a pattern",
			z.strictObject({
				a: z
					.string()
					.regex(/^f\d+$/)
					.describe("a")
			})
		],
		[
			"a default",
			z.strictObject({
				a: z.string().default("x").describe("a")
			})
		],
		[
			"open additional properties",
			z.object({
				a: z.string().describe("a")
			})
		],
		[
			"a float",
			z.strictObject({
				a: z.number().describe("a")
			})
		]
	])("rejects %s", (_label, input) => {
		expect(() =>
			toolInputJsonSchema(
				defineTool({
					name: "probe",
					title: "Probe",
					description: "Probe",
					group: "files",
					effect: "read",
					untrustedOutput: "none",
					exfiltrates: false,
					grounded: {},
					input
				})
			)
		).toThrow(/not portable/)
	})
})

describe("assistant tool registry", () => {
	it("has unique snake_case names", () => {
		const names = ASSISTANT_TOOLS.map(tool => tool.name)

		expect(new Set(names).size).toBe(names.length)

		for (const name of names) {
			expect(name).toMatch(/^[a-z][a-z_]*[a-z]$/)
		}
	})

	it("offers at most six tools per group and every tool in its home group", () => {
		for (const [group, definition] of Object.entries(ASSISTANT_TOOL_GROUPS)) {
			expect(definition.tools.length, group).toBeLessThanOrEqual(6)
		}

		for (const tool of ASSISTANT_TOOLS) {
			expect(ASSISTANT_TOOL_GROUPS[tool.group].tools, tool.name).toContain(tool)
		}
	})

	it("offers a tool under its registry name with the same metadata and at most its arguments", () => {
		for (const [group, definition] of Object.entries(ASSISTANT_TOOL_GROUPS)) {
			for (const offered of definition.tools) {
				const registered = ASSISTANT_TOOLS.find(tool => tool.name === offered.name)
				const properties = Object.keys(toolInputJsonSchema(offered).properties)

				expect(registered, `${group}: ${offered.name}`).toBeDefined()
				expect(portableSchemaViolations(toolInputJsonSchema(offered))).toEqual([])
				expect(Object.keys(toolInputJsonSchema(registered ?? offered).properties)).toEqual(expect.arrayContaining(properties))
				expect([offered.effect, offered.untrustedOutput, offered.exfiltrates, offered.grounded]).toEqual([
					registered?.effect,
					registered?.untrustedOutput,
					registered?.exfiltrates,
					registered?.grounded
				])

				// Small models fill every optional argument, so drive filters come only with the insights route. Notes
				// carry one date and their list is cut at a fixed length, so list_notes filters before the cut wherever
				// it is offered.
				if (group !== "insights" && offered.name !== "list_notes") {
					expect(properties, `${group}: ${offered.name}`).not.toContain("period")
				}

				if (group !== "insights" && offered.name === "list_directory") {
					expect(properties, `${group}: ${offered.name}`).not.toContain("kind")
				}
			}
		}
	})

	it("describes every route", () => {
		for (const route of ASSISTANT_ROUTES) {
			expect(assistantRouteDescription(route).length).toBeGreaterThan(0)
		}
	})

	it("offers no irreversible deletion", () => {
		for (const tool of ASSISTANT_TOOLS) {
			expect(tool.name).not.toMatch(/delete|empty|purge|unshare|leave|account/)
		}

		// The only way to delete is the restorable trash; a new modifying tool has to be added here on purpose.
		expect(
			ASSISTANT_TOOLS.filter(tool => tool.effect === "modifying")
				.map(tool => tool.name)
				.sort()
		).toEqual(["move_items", "rename_item", "set_favorite", "set_note_flags", "trash_items"])
		expect(ASSISTANT_TOOL_GROUPS.trash.tools.filter(isConsequentialTool).map(tool => tool.name)).toEqual(["trash_items"])
	})

	// Pinned by name: deriving the expectation from effect/exfiltrates would restate isConsequentialTool
	// and pass for any registry, e.g. one where draft_chat_message lost its exfiltrates flag.
	it("treats exactly the data-changing and outbound tools as consequential", () => {
		expect(
			ASSISTANT_TOOLS.filter(isConsequentialTool)
				.map(tool => tool.name)
				.sort()
		).toEqual([
			"add_tracks_to_playlist",
			"copy_items",
			"create_directory",
			"create_note",
			"draft_chat_message",
			"move_items",
			"rename_item",
			"set_favorite",
			"set_note_flags",
			"trash_items"
		])
		expect(ASSISTANT_TOOLS.filter(tool => tool.exfiltrates).map(tool => tool.name)).toEqual(["draft_chat_message"])
	})

	it("grounds only declared arguments and every handle of a consequential tool", () => {
		for (const tool of ASSISTANT_TOOLS) {
			const properties = Object.keys(toolInputJsonSchema(tool).properties)
			const grounded = [...(tool.grounded.handles ?? []), ...(tool.grounded.text ?? [])]

			for (const key of grounded) {
				expect(properties, `${tool.name}.${key}`).toContain(key)
			}

			if (isConsequentialTool(tool)) {
				const handleArgs = properties.filter(
					key =>
						key === "handle" ||
						key === "handles" ||
						key === "parent" ||
						key === "destination" ||
						key === "chat" ||
						key === "playlist"
				)

				expect(tool.grounded.handles ?? [], tool.name).toEqual(expect.arrayContaining(handleArgs))
			}
		}
	})

	it("marks every tool that returns other people's text as untrusted", () => {
		const content = ASSISTANT_TOOLS.filter(tool => tool.untrustedOutput === "content").map(tool => tool.name)

		expect(content.sort()).toEqual(["read_chat", "read_note", "read_text_file"])

		for (const tool of ASSISTANT_TOOLS) {
			// audio_control names the track now playing, which may be someone else's file name; get_storage reads
			// only the account's own counters.
			expect(tool.untrustedOutput === "none", tool.name).toBe(tool.name === "open_route" || tool.name === "get_storage")
		}
	})
})
