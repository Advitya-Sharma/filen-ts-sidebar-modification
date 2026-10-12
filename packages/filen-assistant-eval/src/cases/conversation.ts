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

// A note and a chat that share a word ("offsite"), so a request for one right after the other can be pulled
// to the earlier one's route.
function addOffsite(state: AppState): void {
	state.notes.push({
		handle: "n40",
		title: "Offsite shopping list",
		type: "checklist",
		content: "[ ] name badges\n[ ] coffee for 12\n[ ] flip chart paper",
		tags: [],
		pinned: false,
		favorite: false,
		archived: false,
		trashed: false,
		edited: "2025-10-06T10:00:00Z"
	})
	state.chats.push({
		handle: "c40",
		name: "Offsite planning",
		participants: ["Lena", "Tom"],
		messages: [
			{ from: "Lena", text: "The venue is booked: Haus am See, 13 and 14 November.", sent: "2025-10-06T09:00:00Z" },
			{ from: "Tom", text: "We can spend at most 4,800 EUR, venue included.", sent: "2025-10-06T09:30:00Z" },
			{ from: "me", text: "I'll bring the shopping list.", sent: "2025-10-06T10:00:00Z" }
		]
	})
}

// Only the last page holds these.
function lastScans(count: number): string[] {
	return [count, count - 1, count - 2].map(index => `Scan_${String(index).padStart(3, "0")}`)
}

const TWO_PAGES = LISTING_PAGE_ITEMS + 5
const THREE_PAGES = LISTING_PAGE_ITEMS * 2 + 5
const DIRECTORY_COUNT = 214

// A long conversation of requests each answered from a tool, as people use the assistant: later requests still
// call their tools rather than answering from what earlier replies said about a similar item or action.
const LONG_CONVERSATION = [
	"What's in my Cloud Drive root directory?",
	"Find my tax documents.",
	"Which files did I change most recently?",
	"What's my largest file?",
	"What does my note about the offsite shopping list say?",
	"Summarize my offsite planning chat.",
	"Tell the offsite planning chat that I'll bring the shopping list.",
	"What does my note about the launch meeting say?"
]

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
	// Content read after an unrelated turn: the extraction answers this request, and the reply is in its
	// language (graded as a language case), neither answering the earlier turn nor in the earlier turn's language.
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
			// A German turn, then an English one: the answer is English. The excluded German words also catch a
			// reply too short for the language detector.
			{
				id: "de-en",
				lang: "en",
				earlier: "Wie viel Speicherplatz habe ich noch frei?",
				prompt: "What does my note Einkaufsliste say?",
				includes: [
					["milch", "milk"],
					["käse", "cheese"]
				],
				excludes: ["speicher", "storage", "deine", "enthält", " und "]
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
		category: "language",
		lang,
		prompt: earlier,
		followUps: [prompt],
		routes: ["notes"],
		calls: [{ tool: "read_note" }],
		answerIncludes: includes,
		answerExcludes: [...excludes, ...(BENIGN_EXCLUDES[lang] ?? [])],
		expectsText: true
	})),
	// A new request that names another source sharing a word with the previous one: routed by its own words.
	...(
		[
			["en", "What does my note about the offsite shopping list say?", "Summarize my offsite planning chat."],
			["en-now", "What does my note about the offsite shopping list say?", "Now summarize my chat about the offsite."],
			["de", "Was steht in meiner Notiz zur Offsite-Einkaufsliste?", "Fasse meinen Chat zur Offsite-Planung zusammen."],
			["ja", "Offsite shopping listのメモには何が書いてある？", "Offsite planningのチャットを要約して"]
		] as const
	).map(([id, prompt, followUp]): EvalCase => ({
		id: `conversation-switch-note-chat-${id}`,
		category: "multi_step",
		lang: id === "en-now" ? "en" : id,
		prompt,
		followUps: [followUp],
		setup: addOffsite,
		routes: ["chats"],
		calls: [{ tool: "read_chat", args: { handle: "c40" } }],
		answerIncludes: [["4,800", "4.800", "4800", "4 800", "haus am see"]],
		expectsText: true
	})),
	...(
		[
			["en", "Summarize my offsite planning chat.", "What does my note about the offsite shopping list say?"],
			["es", "Resume mi chat de planificación del offsite.", "¿Qué dice mi nota con la lista de compras del offsite?"],
			["fr", "Résume ma discussion sur l'organisation de l'offsite.", "Que dit ma note sur la liste de courses de l'offsite ?"]
		] as const
	).map(([lang, prompt, followUp]): EvalCase => ({
		id: `conversation-switch-chat-note-${lang}`,
		category: "multi_step",
		lang,
		prompt,
		followUps: [followUp],
		setup: addOffsite,
		routes: ["notes"],
		calls: [{ tool: "read_note", args: { handle: "n40" } }],
		answerIncludes: [["12"]],
		expectsText: true
	})),
	...(
		[
			["en", "What does roadmap.md say about the launch?", "And my meeting note, what does it say about the launch?"],
			["pt", "O que diz o roadmap.md sobre o lançamento?", "E a minha nota Meeting 2025-10-01, o que diz sobre o lançamento?"]
		] as const
	).map(([lang, prompt, followUp]): EvalCase => ({
		id: `conversation-switch-file-note-${lang}`,
		category: "multi_step",
		lang,
		prompt,
		followUps: [followUp],
		routes: ["notes"],
		calls: [{ tool: "read_note", args: { handle: "n3" } }],
		answerIncludes: [["15"]],
		expectsText: true
	})),
	...(
		[
			["en", "What does my meeting note say about the launch?", "Now check roadmap.md: what does it say about the launch?"],
			["zh", "我的会议笔记里关于发布说了什么？", "那roadmap.md里关于发布写了什么？"]
		] as const
	).map(([lang, prompt, followUp]): EvalCase => ({
		id: `conversation-switch-note-file-${lang}`,
		category: "multi_step",
		lang,
		prompt,
		followUps: [followUp],
		routes: ["files"],
		calls: [{ tool: "read_text_file", args: { handle: "f27" } }],
		answerIncludes: [["12,000", "12.000", "12000", "12 000", "cookie", "1.2万", "1.2 万"]],
		expectsText: true
	})),
	// The same switch after a reply that asked a question (an offer to read the item it found).
	...(
		[
			["en", "Do I have a note about the offsite shopping list? Ask me before you read it.", "Summarize my offsite planning chat."],
			[
				"de",
				"Habe ich eine Notiz zur Offsite-Einkaufsliste? Frag mich, bevor du sie liest.",
				"Fasse meinen Chat zur Offsite-Planung zusammen."
			],
			["ja", "Offsite shopping listのメモはある？読む前に聞いて。", "Offsite planningのチャットを要約して"]
		] as const
	).map(([lang, prompt, followUp]): EvalCase => ({
		id: `conversation-switch-offer-note-chat-${lang}`,
		category: "multi_step",
		lang,
		prompt,
		followUps: [followUp],
		setup: addOffsite,
		routes: ["chats"],
		calls: [{ tool: "read_chat", args: { handle: "c40" } }],
		answerIncludes: [["4,800", "4.800", "4800", "4 800", "haus am see"]],
		expectsText: true
	})),
	...(
		[
			[
				"en",
				"Do I have a chat about the offsite planning? Ask me before you read it.",
				"What does my note about the offsite shopping list say?"
			],
			[
				"es",
				"¿Tengo un chat sobre la planificación del offsite? Pregúntame antes de leerlo.",
				"¿Qué dice mi nota con la lista de compras del offsite?"
			]
		] as const
	).map(([lang, prompt, followUp]): EvalCase => ({
		id: `conversation-switch-offer-chat-note-${lang}`,
		category: "multi_step",
		lang,
		prompt,
		followUps: [followUp],
		setup: addOffsite,
		routes: ["notes"],
		calls: [{ tool: "read_note", args: { handle: "n40" } }],
		answerIncludes: [["12"]],
		expectsText: true
	})),
	// A short follow-up that picks from what the previous turn listed: routed like that turn.
	...(
		[
			["en", "List my notes", "Read the second one"],
			["de", "Liste meine Notizen auf", "Lies die zweite"],
			["ja", "メモを一覧にして", "2番目を読んで"]
		] as const
	).map(([lang, prompt, followUp]): EvalCase => ({
		id: `conversation-ordinal-note-${lang}`,
		category: "multi_step",
		lang,
		prompt,
		followUps: [followUp],
		routes: ["notes"],
		calls: [{ tool: "read_note" }]
	})),
	{
		id: "conversation-ordinal-chat-en",
		category: "multi_step",
		lang: "en",
		prompt: "List my chats",
		followUps: ["Summarize the first one"],
		routes: ["chats"],
		calls: [{ tool: "read_chat" }]
	},
	{
		id: "conversation-ordinal-file-de",
		category: "multi_step",
		lang: "de",
		prompt: "Was liegt in meinem Ordner Documents?",
		followUps: ["Zeig mir die zweite"],
		routes: ["files"]
	},
	// A directory of many pages: the next page follows, not a count.
	...(
		[
			["en", "What's in my Scans folder?", "What else is in there?"],
			["de", "Was liegt in meinem Ordner Scans?", "Und was ist sonst noch drin?"]
		] as const
	).map(([lang, prompt, followUp]): EvalCase => ({
		id: `conversation-list-page-large-${lang}`,
		category: "multi_step",
		lang,
		prompt,
		followUps: [followUp],
		setup: addScans(DIRECTORY_COUNT),
		routes: ["files"],
		calls: [{ tool: "list_directory", args: { offset: LISTING_PAGE_ITEMS } }],
		answerIncludes: [["Scan_031", "Scan_045", "Scan_060"]],
		expectsText: true
	})),
	// A directory longer than a listing page: the count is the listing's, not rows added up.
	...(
		[
			["en", "How many files are in my Scans folder?"],
			["de", "Wie viele Dateien liegen in meinem Ordner Scans?"],
			["ja", "Scansフォルダにはファイルがいくつありますか？"]
		] as const
	).map(([lang, prompt]): EvalCase => ({
		id: `simple-count-directory-${lang}`,
		category: "simple",
		lang,
		prompt,
		setup: addScans(DIRECTORY_COUNT),
		routes: ["files", "insights"],
		answerIncludes: [[String(DIRECTORY_COUNT)]],
		expectsText: true
	})),
	// Replies that only continue the previous request follow its route.
	...(
		[
			["en", "Do I have a note called Groceries? Ask me before you read it.", "yes"],
			["de", "Habe ich eine Notiz namens Einkaufsliste? Frag mich, bevor du sie liest.", "ja"]
		] as const
	).map(([lang, prompt, followUp]): EvalCase => ({
		id: `conversation-continue-yes-${lang}`,
		category: "multi_step",
		lang,
		prompt,
		followUps: [followUp],
		routes: ["notes"],
		calls: [{ tool: "read_note", args: { handle: lang === "en" ? "n1" : "n2" } }],
		answerIncludes: [lang === "en" ? ["milk", "coffee"] : ["milch", "käse"]],
		expectsText: true
	})),
	...(
		[
			["en", "Find my Telekom invoices and ask me which one to show.", "the second one"],
			["id", "Cari faktur Telekom saya, lalu tanyakan mana yang harus ditampilkan.", "yang kedua"]
		] as const
	).map(([lang, prompt, followUp]): EvalCase => ({
		id: `conversation-continue-choice-${lang}`,
		category: "multi_step",
		lang,
		prompt,
		followUps: [followUp],
		routes: ["files"]
	})),
	{
		id: "conversation-continue-language-de",
		category: "multi_step",
		lang: "de",
		prompt: "Summarize my note Groceries",
		followUps: ["and in German?"],
		routes: ["notes"],
		answerIncludes: [["milch"], ["kaffee", "eier"]],
		expectsText: true
	},
	{
		id: "conversation-continue-language-en",
		category: "multi_step",
		lang: "en",
		prompt: "会議メモを要約して",
		followUps: ["英語で？"],
		routes: ["notes"],
		answerIncludes: [["design"], ["friday"]],
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
	},
	{
		id: "conversation-long-chat-after-note-en",
		category: "multi_step",
		lang: "en",
		prompt: LONG_CONVERSATION[0] ?? "",
		followUps: [...LONG_CONVERSATION.slice(1), "Summarize my website team chat about the launch."],
		setup: addOffsite,
		routes: ["chats"],
		calls: [{ tool: "read_chat", args: { handle: "c4" } }],
		answerIncludes: [["review", "draft"]],
		// What the launch note says, not the chat.
		answerExcludes: ["8 october", "october 8", "owns", "responsible"],
		// The draft turn's.
		state: [{ draft: "c40" }],
		expectsText: true
	},
	// The same with reads only: a storage question in place of the draft.
	{
		id: "conversation-long-reads-chat-after-note-en",
		category: "multi_step",
		lang: "en",
		prompt: LONG_CONVERSATION[0] ?? "",
		followUps: [
			...LONG_CONVERSATION.slice(1, 6),
			"How much storage am I using?",
			LONG_CONVERSATION[7] ?? "",
			"Summarize my website team chat about the launch."
		],
		setup: addOffsite,
		routes: ["chats"],
		calls: [{ tool: "read_chat", args: { handle: "c4" } }],
		answerIncludes: [["review", "draft"]],
		answerExcludes: ["8 october", "october 8", "owns", "responsible"],
		expectsText: true
	},
	{
		id: "conversation-long-open-after-open-en",
		category: "multi_step",
		lang: "en",
		prompt: LONG_CONVERSATION[0] ?? "",
		followUps: [...LONG_CONVERSATION.slice(1), "Take me to my favorites", "Open my trash"],
		setup: addOffsite,
		routes: ["navigation", "trash"],
		calls: [{ tool: "open_route", args: { route: "trash" } }],
		state: [{ screen: "trash" }, { draft: "c40" }]
	}
]
