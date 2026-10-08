import type { AppState } from "../fixture"
import type { EvalCase } from "./types"

// The eval's clock (prompts.ts) is Wednesday 2025-10-08 in Europe/Berlin: yesterday 2025-10-07, this week
// 2025-10-06 to 2025-10-12, last week 2025-09-29 to 2025-10-05. Fixture notes default to 2025-09-20 and
// chat messages to 2025-10-01..03.

function editNotes(state: AppState, edited: Readonly<Record<string, string>>, others?: string): void {
	for (const note of state.notes) {
		const date = edited[note.handle] ?? others

		if (date !== undefined) {
			note.edited = date
		}
	}
}

function sendMessages(state: AppState, chat: string, sent: readonly string[]): void {
	const messages = state.chats.find(item => item.handle === chat)?.messages ?? []

	sent.forEach((date, index) => {
		const message = messages[index]

		if (message !== undefined) {
			message.sent = date
		}
	})
}

// Today's date, weekday, month or year, from the prompt's date line alone.
const CALENDAR: (Pick<EvalCase, "lang" | "prompt" | "answerIncludes" | "answerExcludes"> & { topic: string })[] = [
	{
		topic: "today-date",
		lang: "en",
		prompt: "What's today's date?",
		answerIncludes: [
			["8", "2025-10-08"],
			["october", "2025-10-08", "10/8"]
		]
	},
	{
		topic: "weekday",
		lang: "es",
		prompt: "¿Qué día de la semana es hoy?",
		answerIncludes: [["miercoles"]]
	},
	{
		topic: "month",
		lang: "de",
		prompt: "Welcher Monat ist gerade?",
		answerIncludes: [["oktober"]],
		answerExcludes: ["juli"]
	},
	{
		topic: "today-date",
		lang: "ja",
		prompt: "今日は何日？",
		answerIncludes: [["8", "2025-10-08"]]
	},
	{
		topic: "yesterday",
		lang: "pt",
		prompt: "Que dia foi ontem?",
		answerIncludes: [["7", "terca"]]
	},
	{
		topic: "this-week",
		lang: "fr",
		prompt: "Du combien au combien va cette semaine ?",
		answerIncludes: [
			["6", "2025-10-06"],
			["12", "2025-10-12"]
		],
		answerExcludes: ["vendredi", "mardi"]
	},
	{
		topic: "last-month",
		lang: "ru",
		prompt: "Какой месяц был прошлым?",
		answerIncludes: [["сентябр"]]
	},
	{
		topic: "year",
		lang: "zh",
		prompt: "今年是哪一年？",
		answerIncludes: [["2025"]]
	},
	{
		topic: "today-date",
		lang: "ar",
		prompt: "ما هو تاريخ اليوم؟",
		answerIncludes: [
			["8", "٨", "2025-10-08"],
			["أكتوبر", "اكتوبر", "10", "١٠", "2025-10-08"]
		]
	},
	{
		topic: "weekday",
		lang: "hi",
		prompt: "आज कौन सा दिन है?",
		answerIncludes: [["बुधवार", "wednesday"]]
	},
	{
		topic: "weekday",
		lang: "id",
		prompt: "Hari ini hari apa?",
		answerIncludes: [["rabu"]]
	},
	{
		topic: "today-date",
		lang: "tr",
		prompt: "Bugün ayın kaçı?",
		answerIncludes: [["8", "2025-10-08"]]
	}
]

export const DATE_CASES: EvalCase[] = [
	...CALENDAR.map(({ topic, ...item }): EvalCase => ({
		...item,
		id: `dates-${topic}-${item.lang}`,
		category: "dates",
		routes: ["answer"],
		expectsText: true
	})),
	{
		id: "dates-notes-this-week-en",
		category: "dates",
		lang: "en",
		prompt: "Which notes did I edit this week?",
		setup: state =>
			editNotes(state, {
				n1: "2025-10-06T18:00:00Z",
				n3: "2025-10-07T09:00:00Z"
			}),
		routes: ["notes"],
		calls: [
			{
				tool: "list_notes"
			}
		],
		answerIncludes: [["groceries"], ["meeting"]],
		answerExcludes: ["einkaufsliste", "tortilla", "Важное"],
		expectsText: true
	},
	{
		id: "dates-notes-last-month-zh",
		category: "dates",
		lang: "zh",
		prompt: "我上个月编辑过哪些笔记？",
		setup: state =>
			editNotes(
				state,
				{
					n8: "2025-09-12T10:00:00Z",
					n9: "2025-10-02T10:00:00Z"
				},
				"2025-08-15T10:00:00Z"
			),
		routes: ["notes"],
		calls: [
			{
				tool: "list_notes"
			}
		],
		answerIncludes: [["读书笔记"]],
		answerExcludes: ["会議メモ", "groceries", "einkaufsliste"],
		expectsText: true
	},
	{
		id: "dates-chat-yesterday-de",
		category: "dates",
		lang: "de",
		prompt: "Was hat Anna mir gestern geschrieben?",
		setup: state => sendMessages(state, "c1", ["2025-10-01T10:00:00Z", "2025-10-02T11:00:00Z", "2025-10-07T15:00:00Z"]),
		routes: ["chats"],
		calls: [
			{
				tool: "read_chat",
				args: {
					handle: "c1"
				}
			}
		],
		answerIncludes: [["freitag", "termin"]],
		answerExcludes: ["hetzner"],
		expectsText: true
	},
	{
		id: "dates-chat-today-ja",
		category: "dates",
		lang: "ja",
		prompt: "今日Yukiから何か来た？",
		setup: state => sendMessages(state, "c5", ["2025-10-08T06:00:00Z", "2025-10-08T06:30:00Z"]),
		routes: ["chats"],
		calls: [
			{
				tool: "read_chat",
				args: {
					handle: "c5"
				}
			}
		],
		answerIncludes: [["金曜"]],
		expectsText: true
	},
	{
		id: "dates-chats-this-week-es",
		category: "dates",
		lang: "es",
		prompt: "¿Qué chats tuvieron mensajes esta semana?",
		setup: state => {
			sendMessages(state, "c3", ["2025-10-01T10:00:00Z", "2025-10-07T19:00:00Z"])
			sendMessages(state, "c5", ["2025-10-02T10:00:00Z", "2025-10-06T20:00:00Z"])
		},
		routes: ["chats"],
		calls: [
			{
				tool: "list_chats"
			}
		],
		answerIncludes: [["carlos"], ["yuki"]],
		answerExcludes: ["anna", "family", "website"],
		expectsText: true
	},
	{
		id: "dates-file-this-month-fr",
		category: "dates",
		lang: "fr",
		prompt: "Est-ce que roadmap.md a été modifié ce mois-ci ?",
		routes: ["files"],
		calls: [
			{
				tool: "get_item_info",
				args: {
					handle: "f27"
				}
			}
		],
		answerIncludes: [["septembre", "29", "2025-09-29", "pas ete modifie"]],
		answerExcludes: ["oui"],
		expectsText: true
	},
	{
		id: "dates-file-last-week-ru",
		category: "dates",
		lang: "ru",
		prompt: "Файл report_Q3.pdf менялся на прошлой неделе?",
		routes: ["files"],
		calls: [
			{
				tool: "get_item_info",
				args: {
					handle: "f33"
				}
			}
		],
		answerIncludes: [["октябр", "2025-10-01", "01.10"]],
		answerExcludes: ["нет"],
		expectsText: true
	},
	// Recents (this month: f33, f36 and the old photo f14 uploaded 10-06; last week: f27, f28, f33, f36), filtered
	// by a period the code resolves.
	{
		id: "dates-drive-last-week-de",
		category: "dates",
		lang: "de",
		prompt: "Welche Dateien habe ich letzte Woche geändert?",
		routes: ["insights"],
		calls: [
			{
				tool: "recent_items"
			}
		],
		answerIncludes: [["roadmap"], ["meeting_notes"], ["report_q3"], ["報告書"]],
		// This week's range, which a model without a last-week range names instead.
		answerExcludes: ["2025-10-06", "06.10", "2025-10-12", "12.10", "setup_installer", "oma_geburtstag"],
		expectsText: true
	},
	{
		id: "dates-drive-this-month-en",
		category: "dates",
		lang: "en",
		prompt: "Which files did I upload this month?",
		routes: ["insights"],
		calls: [
			{
				tool: "recent_items"
			}
		],
		answerIncludes: [["oma_geburtstag"], ["report_q3"], ["報告書"]],
		answerExcludes: ["setup_installer", "tugas_kuliah", "receipt_amazon", "roadmap", "meeting_notes", "logo_final"],
		expectsText: true
	},
	{
		id: "dates-drive-yesterday-pt",
		category: "dates",
		lang: "pt",
		prompt: "Quais arquivos eu modifiquei ontem?",
		routes: ["insights"],
		calls: [
			{
				tool: "recent_items"
			}
		],
		answerExcludes: [
			"report_Q3",
			"roadmap",
			"meeting_notes",
			"setup_installer",
			"trabalho_de_historia",
			"redação_enem",
			"oma_geburtstag"
		],
		expectsText: true
	}
]
