import { LISTING_PAGE_ITEMS } from "@filen/shared/assistant"
import type { AppState } from "../fixture"
import { BENIGN_READ_EXCLUDES } from "./single"
import type { EvalCase, Language } from "./types"

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

const BENIGN_EXCLUDES: Partial<Record<Language, readonly string[]>> = BENIGN_READ_EXCLUDES

// Two contacts called Tom with a chat each, so "Tom" alone is ambiguous; Tom Lee's chat is long.
function addTomChats(lee: readonly (readonly [string, string])[]): (state: AppState) => void {
	return state => {
		state.chats.push(
			{
				handle: "c30",
				name: "Tom Lee",
				participants: ["Tom Lee"],
				messages: lee.map(([from, text], index) => ({
					from,
					text,
					sent: `2025-10-0${String(1 + Math.floor(index / 3))}T10:00:00Z`
				}))
			},
			{
				handle: "c31",
				name: "Tom Smith",
				participants: ["Tom Smith"],
				messages: [
					{ from: "Tom Smith", text: "Can I borrow your drill this weekend?", sent: "2025-10-04T09:00:00Z" },
					{ from: "me", text: "Sure, pick it up on Friday.", sent: "2025-10-04T09:30:00Z" }
				]
			}
		)
		state.contacts.push(
			{ name: "Tom Lee", email: "tom.lee@example.com", chat: "c30" },
			{ name: "Tom Smith", email: "tom.smith@example.com", chat: "c31" }
		)
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
	// Content read after an unrelated turn: the extraction answers this request, in its language, and neither
	// the earlier one nor about it.
	...(
		[
			{
				id: "en",
				lang: "en",
				earlier: "How much storage do I have left?",
				prompt: "Summarize my note Groceries",
				includes: [["milk"], ["coffee", "eggs"]],
				excludes: ["storage"]
			},
			// A German turn, then an English one: the answer is English.
			{
				id: "de-en",
				lang: "en",
				earlier: "Wie viel Speicherplatz habe ich noch frei?",
				prompt: "What does my note Einkaufsliste say?",
				includes: [
					["milch", "milk"],
					["käse", "cheese"]
				],
				excludes: ["speicher", "storage"]
			},
			{
				id: "es",
				lang: "es",
				earlier: "Pon mi playlist Workout",
				prompt: "¿Qué dice mi nota Receta tortilla de patatas?",
				includes: [["huevos"], ["35"]],
				excludes: ["workout", "playlist"]
			},
			{
				id: "tr",
				lang: "tr",
				earlier: "Geçen hafta benimle kim dosya paylaştı?",
				prompt: "Alışveriş listesi notumda ne yazıyor?",
				includes: [
					["süt", "sut"],
					["ekmek", "domates"]
				],
				excludes: ["paylaş"]
			},
			{
				id: "fr",
				lang: "fr",
				earlier: "Quel temps fera-t-il demain à Berlin ?",
				prompt: "Résume ma note Idées cadeaux Noël",
				includes: [["écharpe", "echarpe"], ["casque"]],
				excludes: ["berlin", "météo", "meteo"]
			},
			// Requests as short as a reply to a clarifying question.
			{
				id: "ja",
				lang: "ja",
				earlier: "ワークアウトのプレイリストを再生して",
				prompt: "会議メモには何て書いてある？",
				includes: [["金曜"], ["デザイン"]],
				excludes: ["ワークアウト", "プレイリスト", "再生"]
			},
			{
				id: "zh",
				lang: "zh",
				earlier: "我还剩多少存储空间？",
				prompt: "我的读书笔记写了什么？",
				includes: [["三体"], ["活着"]],
				excludes: ["存储", "空间"]
			}
		] as const
	).map(({ id, lang, earlier, prompt, includes, excludes }): EvalCase => ({
		id: `conversation-read-after-unrelated-${id}`,
		category: "multi_step",
		lang,
		prompt: earlier,
		followUps: [prompt],
		routes: ["notes"],
		calls: [{ tool: "read_note" }],
		answerIncludes: includes,
		answerExcludes: [...excludes, ...(BENIGN_EXCLUDES[lang] ?? [])],
		expectsText: true
	})),
	{
		id: "conversation-clarified-extraction-de",
		category: "multi_step",
		lang: "de",
		prompt: "Was schreibt Anna über den Termin?",
		followUps: ["Anna Schmidt"],
		routes: ["chats"],
		answerIncludes: [["freitag"]],
		expectsText: true
	},
	// The reply to a clarifying question names only the chat; the extraction must still look for what the
	// earlier message asked, which a general summary of a long chat leaves out. The first turn asks for the
	// question, which the model otherwise skips by reading every chat with a Tom.
	{
		id: "conversation-clarified-long-chat-en",
		category: "multi_step",
		lang: "en",
		prompt: "What did Tom write about the offsite budget? I have chats with several Toms, ask me which one first.",
		followUps: ["Tom Lee"],
		setup: addTomChats([
			["Tom Lee", "Are you coming to the climbing gym on Thursday?"],
			["me", "Yes, around 7."],
			["Tom Lee", "Great, I'll book the wall for two."],
			["Tom Lee", "Did you see the new coffee place on Main Street?"],
			["me", "Not yet, is it good?"],
			["Tom Lee", "Best flat white in town, honestly."],
			["Tom Lee", "For the offsite we can spend at most 4,800 euros, venue included."],
			["me", "Got it."],
			["Tom Lee", "My sister is visiting next weekend, so I'm off on Saturday."],
			["me", "Have fun!"],
			["Tom Lee", "Can you send me the photos from the hike?"],
			["me", "Uploading them tonight."],
			["Tom Lee", "The car is back from the garage, finally."],
			["Tom Lee", "I switched to the new phone plan too, much cheaper."],
			["me", "Nice, which one?"],
			["Tom Lee", "The one with 50 GB. See you Thursday!"]
		]),
		routes: ["chats"],
		answerIncludes: [["4,800", "4800", "4.800", "4 800"]],
		expectsText: true
	},
	{
		id: "conversation-clarified-long-chat-de",
		category: "multi_step",
		lang: "de",
		prompt: "Was hat Tom zum Budget für das Offsite geschrieben? Ich habe Chats mit mehreren Toms, frag mich zuerst, welchen.",
		followUps: ["Tom Lee"],
		setup: addTomChats([
			["Tom Lee", "Kommst du Donnerstag mit in die Kletterhalle?"],
			["me", "Ja, so gegen 7."],
			["Tom Lee", "Super, ich reserviere die Wand für zwei."],
			["Tom Lee", "Hast du das neue Café in der Hauptstraße gesehen?"],
			["me", "Noch nicht, ist es gut?"],
			["Tom Lee", "Der beste Flat White der Stadt, ehrlich."],
			["Tom Lee", "Fürs Offsite können wir höchstens 4.800 Euro ausgeben, Location inklusive."],
			["me", "Alles klar."],
			["Tom Lee", "Meine Schwester kommt nächstes Wochenende, Samstag bin ich weg."],
			["me", "Viel Spaß!"],
			["Tom Lee", "Schickst du mir die Fotos von der Wanderung?"],
			["me", "Lade sie heute Abend hoch."],
			["Tom Lee", "Das Auto ist endlich aus der Werkstatt zurück."],
			["Tom Lee", "Ich hab auch zum neuen Handytarif gewechselt, viel günstiger."],
			["me", "Cool, welcher?"],
			["Tom Lee", "Der mit 50 GB. Bis Donnerstag!"]
		]),
		routes: ["chats"],
		answerIncludes: [["4.800", "4800", "4,800", "4 800"]],
		expectsText: true
	}
]
