import type { EvalCase } from "./types"

// Several turns of one conversation: the answer to a clarifying question, a follow-up that only makes sense
// with the turn before, and a request that names the item a previous turn listed.
export const CONVERSATION_CASES: readonly EvalCase[] = [
	{
		id: "conversation-draft-clarified-en",
		category: "multi_step",
		lang: "en",
		prompt: "Write a message to Anna Schmidt",
		followUps: ["That I paid the Hetzner invoice"],
		routes: ["chats"],
		state: [{ draft: "c1", textIncludes: ["hetzner"] }],
		expectsText: true
	},
	{
		id: "conversation-list-followup-en",
		category: "multi_step",
		lang: "en",
		prompt: "What's in my Documents folder?",
		followUps: ["And in Invoices?"],
		routes: ["files"],
		calls: [{ tool: "list_directory" }],
		answerIncludes: [["Telekom", "Hetzner", "Stadtwerke"]],
		expectsText: true
	},
	{
		id: "conversation-list-followup-de",
		category: "multi_step",
		lang: "de",
		prompt: "Was liegt in meinem Ordner Documents?",
		followUps: ["Und in Verträge?"],
		routes: ["files"],
		calls: [{ tool: "list_directory" }],
		answerIncludes: [["Mietvertrag", "Arbeitsvertrag"]],
		expectsText: true
	}
]
