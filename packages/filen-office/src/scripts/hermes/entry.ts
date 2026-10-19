import { decodeBase64, encodeBase64 } from "./base64"
import { INPUT_GLOBAL, RESULT_PREFIX } from "./protocol"
import { runScenario } from "./scenario"

// The bundle's main module. The runner prepends a script that sets the input global; the result goes out as
// one line through the host's print (the Hermes CLI's, or the Node runner's stand-in).

interface Input {
	readonly fixtures: readonly { readonly name: string; readonly base64: string }[]
}

function isInput(value: unknown): value is Input {
	return typeof value === "object" && value !== null && Array.isArray(Reflect.get(value, "fixtures"))
}

function output(line: string): void {
	const print: unknown = Reflect.get(globalThis, "print")

	if (typeof print === "function") print(line)
}

const input: unknown = Reflect.get(globalThis, INPUT_GLOBAL)

if (isInput(input)) {
	runScenario(input.fixtures.map(fixture => ({ name: fixture.name, bytes: decodeBase64(fixture.base64) }))).then(
		result =>
			output(
				RESULT_PREFIX +
					JSON.stringify({
						...result,
						fixtures: result.fixtures.map(fixture => ({
							...fixture,
							unedited: encodeBase64(fixture.unedited),
							edited: encodeBase64(fixture.edited)
						}))
					})
			),
		(error: unknown) =>
			output(`${RESULT_PREFIX}${JSON.stringify({ fatal: error instanceof Error ? (error.stack ?? error.message) : String(error) })}`)
	)
} else {
	output(`${RESULT_PREFIX}${JSON.stringify({ fatal: "no input" })}`)
}
