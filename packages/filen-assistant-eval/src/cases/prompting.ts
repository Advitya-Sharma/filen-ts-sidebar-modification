import type { EvalCase } from "./types"

// What the instructions ask of replies beyond calling the right tool: say when nothing matched instead of
// opening something close, count what a result left out, ask which one when the user means one of several.
export const PROMPTING_CASES: EvalCase[] = [
	{
		id: "simple-note-not-found-en",
		category: "simple",
		lang: "en",
		prompt: "Read my note about the Bali trip",
		routes: ["notes"],
		calls: [
			{
				tool: "list_notes"
			}
		],
		forbidden: ["read_note", "open_note"],
		answerExcludes: ["dubai", "jaipur"],
		expectsText: true
	},
	{
		id: "simple-file-not-found-es",
		category: "simple",
		lang: "es",
		prompt: "Muéstrame el escaneo de mi pasaporte",
		routes: ["files"],
		calls: [
			{
				tool: "search_items"
			}
		],
		forbidden: ["reveal_item", "read_text_file"],
		answerIncludes: [["no encontr", "no se encontr", "ningun", "no hay", "no tienes", "no pude", "no existe"]],
		expectsText: true
	},
	{
		id: "simple-count-partial-en",
		category: "simple",
		lang: "en",
		prompt: "How many PDF files do I have?",
		routes: ["files"],
		calls: [
			{
				tool: "search_items"
			}
		],
		answerIncludes: [["14"]],
		expectsText: true
	},
	{
		id: "simple-count-partial-de",
		category: "simple",
		lang: "de",
		prompt: "Wie viele PDF-Dateien habe ich?",
		routes: ["files"],
		calls: [
			{
				tool: "search_items"
			}
		],
		answerIncludes: [["14"]],
		expectsText: true
	},
	{
		id: "missing-ambiguous-invoice-de",
		category: "missing_param",
		lang: "de",
		prompt: "Zeig mir die Telekom-Rechnung",
		routes: ["files"],
		forbidden: ["reveal_item"],
		expectsText: true
	}
]
