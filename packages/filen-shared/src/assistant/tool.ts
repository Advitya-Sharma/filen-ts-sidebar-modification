import * as z from "zod"

// zod probes `new Function` when it builds its first object schema; the app's CSP blocks that and reports a
// violation on every page load. Under that CSP the interpreted parser runs anyway. Must run before any schema exists.
z.config({ jitless: true })

// What a call does to the user's data. "navigate" only changes what the app shows; "additive" and
// "modifying" change stored data and are executed only after the user confirms.
export type AssistantToolEffect = "read" | "navigate" | "additive" | "modifying"

export type AssistantToolGroup = "files" | "insights" | "shared" | "organize" | "trash" | "notes" | "chats" | "audio" | "navigation"

// "answer" routes a request to a plain text reply without any tools.
export type AssistantRoute = AssistantToolGroup | "answer"

// Who can author the text a tool returns. Item, chat and playlist names are attacker-controlled
// whenever another user shares the item or writes in the chat; "content" (note bodies, file text,
// chat messages) carries long free text and is the main prompt-injection carrier.
export type AssistantUntrustedOutput = "none" | "names" | "content"

export type AssistantGroundedArgs = {
	// Arguments holding handles (string or string[]) that must trace back to the user's request.
	readonly handles?: readonly string[]
	// Free-text arguments that must appear in the user's own message.
	readonly text?: readonly string[]
	// Free-text arguments that may say anything, but carry no link or address the user did not type.
	readonly links?: readonly string[]
}

export type AssistantTool<Name extends string = string, Input extends z.ZodObject = z.ZodObject> = {
	readonly name: Name
	readonly title: string
	readonly description: string
	readonly group: AssistantToolGroup
	readonly effect: AssistantToolEffect
	readonly input: Input
	readonly untrustedOutput: AssistantUntrustedOutput
	// Moves the user's data, or words about it, to other people.
	readonly exfiltrates: boolean
	readonly grounded: AssistantGroundedArgs
}

export function defineTool<const Name extends string, Input extends z.ZodObject>(
	tool: AssistantTool<Name, Input>
): AssistantTool<Name, Input> {
	return tool
}

// Consequential calls change stored data or reach other people. The dispatcher runs them only
// after the guard (guard.ts) has grounded them and the user has confirmed a card that names
// exactly what happens.
export function isConsequentialTool(tool: AssistantTool): boolean {
	return tool.effect === "additive" || tool.effect === "modifying" || tool.exfiltrates
}

export type PortableScalarSchema =
	| {
			type: "string" | "integer" | "boolean"
			description?: string
	  }
	| {
			type: "string"
			enum: string[]
			description?: string
	  }

export type PortableArraySchema = {
	type: "array"
	items: PortableScalarSchema
	minItems?: number
	description?: string
}

// The subset every target grammar engine handles identically: llama.cpp's json-schema-to-grammar
// (wllama, llama.rn) and node-llama-cpp's GBNF builder diverge on $ref, combinators, number
// bounds and formats.
export type PortableObjectSchema = {
	type: "object"
	properties: Record<string, PortableScalarSchema | PortableArraySchema>
	required?: string[]
	additionalProperties: false
	description?: string
}

const SCALAR_KEYS = new Set(["type", "description", "enum"])
const ARRAY_KEYS = new Set(["type", "description", "items", "minItems"])
const OBJECT_KEYS = new Set(["type", "description", "properties", "required", "additionalProperties"])
const SCALAR_TYPES = new Set(["string", "integer", "boolean"])

function isRecord(value: unknown): value is Record<string, unknown> {
	return typeof value === "object" && value !== null && !Array.isArray(value)
}

function scalarViolations({ schema, path }: { schema: unknown; path: string }): string[] {
	if (!isRecord(schema)) {
		return [`${path}: not an object`]
	}

	const violations = Object.keys(schema)
		.filter(key => !SCALAR_KEYS.has(key))
		.map(key => `${path}: unsupported keyword "${key}"`)

	if (typeof schema["type"] !== "string" || !SCALAR_TYPES.has(schema["type"])) {
		violations.push(`${path}: type must be string, integer or boolean`)
	}

	if ("enum" in schema) {
		const values = schema["enum"]

		if (
			schema["type"] !== "string" ||
			!Array.isArray(values) ||
			values.length === 0 ||
			values.some(value => typeof value !== "string")
		) {
			violations.push(`${path}: enum must be a non-empty list of strings on a string type`)
		}
	}

	return violations
}

function propertyViolations({ schema, path }: { schema: unknown; path: string }): string[] {
	if (isRecord(schema) && schema["type"] === "array") {
		const violations = Object.keys(schema)
			.filter(key => !ARRAY_KEYS.has(key))
			.map(key => `${path}: unsupported keyword "${key}"`)

		if ("minItems" in schema && (!Number.isInteger(schema["minItems"]) || Number(schema["minItems"]) < 0)) {
			violations.push(`${path}: minItems must be a non-negative integer`)
		}

		return [
			...violations,
			...scalarViolations({
				schema: schema["items"],
				path: `${path}.items`
			})
		]
	}

	return scalarViolations({
		schema,
		path
	})
}

export function portableSchemaViolations(schema: unknown): string[] {
	if (!isRecord(schema) || schema["type"] !== "object") {
		return ["root: must be an object schema"]
	}

	const violations = Object.keys(schema)
		.filter(key => !OBJECT_KEYS.has(key))
		.map(key => `root: unsupported keyword "${key}"`)

	if (schema["additionalProperties"] !== false) {
		violations.push("root: additionalProperties must be false")
	}

	const properties = schema["properties"]

	if (!isRecord(properties)) {
		return [...violations, "root: properties missing"]
	}

	for (const [name, property] of Object.entries(properties)) {
		violations.push(
			...propertyViolations({
				schema: property,
				path: name
			})
		)
	}

	const required = schema["required"] ?? []

	if (!Array.isArray(required) || required.some(name => typeof name !== "string" || !(name in properties))) {
		violations.push("root: required must list declared properties")
	}

	return violations
}

function stripSafeIntegerBounds(schema: unknown): void {
	if (!isRecord(schema)) {
		return
	}

	// zod emits the safe-integer range for every z.int(); it carries no information.
	if (schema["minimum"] === Number.MIN_SAFE_INTEGER && schema["maximum"] === Number.MAX_SAFE_INTEGER) {
		delete schema["minimum"]
		delete schema["maximum"]
	}

	for (const value of Object.values(schema)) {
		stripSafeIntegerBounds(value)
	}
}

export function toolInputJsonSchema(tool: AssistantTool): PortableObjectSchema {
	const schema: Record<string, unknown> = z.toJSONSchema(tool.input, {
		io: "input",
		unrepresentable: "throw"
	})

	delete schema["$schema"]
	stripSafeIntegerBounds(schema)

	const violations = portableSchemaViolations(schema)

	if (violations.length > 0) {
		throw new Error(`Tool "${tool.name}" input is not portable: ${violations.join("; ")}`)
	}

	return schema as PortableObjectSchema
}
