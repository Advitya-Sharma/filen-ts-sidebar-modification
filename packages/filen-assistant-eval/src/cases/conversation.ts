import { LISTING_PAGE_ITEMS } from "@filen/shared/assistant"
import type { AppState } from "../fixture"
import type { EvalCase } from "./types"

// A directory one listing page cannot show: the rest takes more calls, each with the offset the last gave.
function addScans(count: number): (state: AppState) => void {
	return state => {
		state.drive.push({
			handle: "d60",
			kind: "directory",
			name: "Scans",
			parent: "d0",
			size: 0,
			modified: "2021-03-01T10:00:00Z",
			favorite: false,
			trashed: false,
			shared: false
		})

		for (let index = 1; index <= count; index++) {
			state.drive.push({
				handle: `f${String(300 + index)}`,
				kind: "file",
				name: `Scan_${String(index).padStart(3, "0")}.pdf`,
				parent: "d60",
				size: 300_000,
				modified: "2021-03-01T10:00:00Z",
				favorite: false,
				trashed: false,
				shared: false
			})
		}
	}
}

// Only the last page holds these.
function lastScans(count: number): string[] {
	return [count, count - 1, count - 2].map(index => `Scan_${String(index).padStart(3, "0")}`)
}

const TWO_PAGES = LISTING_PAGE_ITEMS + 5
const THREE_PAGES = LISTING_PAGE_ITEMS * 2 + 5

// Several turns of one conversation: the answer to a clarifying question, a follow-up that only makes sense
// with the turn before, a request that names the item a previous turn listed, and turns that move to another
// tool group, to one whose tools the prompt shows for another family, or to no tools at all.
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
	},
	...(
		[
			["en", "What's in my Scans folder?", "What else is in there?"],
			["de", "Was liegt in meinem Ordner Scans?", "Und was ist sonst noch drin?"],
			["ja", "Scansフォルダには何が入っていますか？", "他には何がありますか？"]
		] as const
	).map(([lang, prompt, followUp]): EvalCase => ({
		id: `conversation-list-page-${lang}`,
		category: "multi_step",
		lang,
		prompt,
		followUps: [followUp],
		setup: addScans(TWO_PAGES),
		routes: ["files"],
		calls: [{ tool: "list_directory", args: { offset: LISTING_PAGE_ITEMS } }],
		answerIncludes: [lastScans(TWO_PAGES)],
		expectsText: true
	})),
	{
		id: "conversation-same-family-en",
		category: "multi_step",
		lang: "en",
		prompt: "What's in my Documents folder?",
		followUps: ["Is Steuererklärung_2024.pdf in there shared with anyone?"],
		routes: ["shared"],
		calls: [{ tool: "get_sharing", args: { handle: "f10" } }],
		answerIncludes: [["Anna", "anna.schmidt"]],
		expectsText: true
	},
	...(
		[
			["en", "What's in my Documents folder?", "What did Anna write to me?"],
			["ja", "Documentsフォルダには何が入っていますか？", "Annaは何と書いてきましたか？"]
		] as const
	).map(([lang, prompt, followUp]): EvalCase => ({
		id: `conversation-cross-family-${lang}`,
		category: "multi_step",
		lang,
		prompt,
		followUps: [followUp],
		routes: ["chats"],
		calls: [{ tool: "read_chat", args: { handle: "c1" } }],
		answerIncludes: [["Hetzner", "10"]],
		expectsText: true
	})),
	...(
		[
			["en", "What's in my Documents folder?", "Thanks! How do you say thank you in Spanish?"],
			["de", "Was liegt in meinem Ordner Documents?", "Danke! Was heißt danke auf Spanisch?"]
		] as const
	).map(([lang, prompt, followUp]): EvalCase => ({
		id: `conversation-answer-after-tools-${lang}`,
		category: "multi_step",
		lang,
		prompt,
		followUps: [followUp],
		routes: ["answer"],
		answerIncludes: [["gracias"]],
		expectsText: true
	})),
	{
		id: "conversation-list-pages-en",
		category: "multi_step",
		lang: "en",
		prompt: "What's in my Scans folder?",
		followUps: ["What else is in there?", "And the rest?"],
		setup: addScans(THREE_PAGES),
		routes: ["files"],
		calls: [{ tool: "list_directory", args: { offset: LISTING_PAGE_ITEMS * 2 } }],
		answerIncludes: [lastScans(THREE_PAGES)],
		expectsText: true
	},
	// A follow-up about content the previous turn read: the extraction answers the new question, not the old one.
	{
		id: "conversation-note-followup-fr",
		category: "multi_step",
		lang: "fr",
		prompt: "Résume ma note Idées cadeaux Noël",
		followUps: ["Et pour Léa, j'avais noté quoi ?"],
		routes: ["notes"],
		answerIncludes: [["casque"]],
		expectsText: true
	},
	{
		id: "conversation-chat-followup-es",
		category: "multi_step",
		lang: "es",
		prompt: "Lee mi chat con Carlos Ruiz",
		followUps: ["¿Qué me preguntó?"],
		routes: ["chats"],
		answerIncludes: [["presupuesto"]],
		expectsText: true
	},
	{
		id: "conversation-file-followup-en",
		category: "multi_step",
		lang: "en",
		prompt: "What's in tax_notes_2024.txt?",
		followUps: ["When is the appointment with the advisor?"],
		routes: ["files"],
		answerIncludes: [["12 may", "may 12"]],
		expectsText: true
	},
	{
		id: "conversation-clarified-extraction-de",
		category: "multi_step",
		lang: "de",
		prompt: "Was schreibt Anna über den Termin?",
		followUps: ["Anna Schmidt"],
		routes: ["chats"],
		answerIncludes: [["freitag"]],
		expectsText: true
	}
]
