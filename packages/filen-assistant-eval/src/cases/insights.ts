import type { AppState } from "../fixture"
import type { EvalCase } from "./types"

// The eval's clock (prompts.ts) is Wednesday 2025-10-08 in Europe/Berlin. The Recents list holds files
// uploaded since 2025-08-24: this month f33 (10-01), f36 (10-02) and the old photo f14 (taken 2024, uploaded
// 10-06). The account uses 3.2 of 10 GiB (6.8 GiB free, 32%).

function setEdited(state: AppState, edited: Readonly<Record<string, string>>): void {
	for (const note of state.notes) {
		note.edited = edited[note.handle] ?? note.edited
	}
}

// More notes than a listing shows (40), all edited this month, and the two edited last month: one of them
// archived, which the notes list puts last.
function manyNotes(state: AppState): void {
	for (const note of state.notes) {
		note.edited = "2025-10-03T09:00:00Z"
	}

	setEdited(state, {
		n4: "2025-09-15T18:00:00Z",
		n10: "2025-09-28T09:00:00Z"
	})

	const archived = state.notes.find(note => note.handle === "n4")

	if (archived !== undefined) {
		archived.archived = true
	}

	for (let index = 0; index < 45; index++) {
		state.notes.push({
			handle: `n${String(100 + index)}`,
			title: `Daily log ${String(index + 1)}`,
			type: "text",
			content: "",
			tags: [],
			pinned: false,
			favorite: false,
			archived: false,
			trashed: false,
			edited: `2025-10-0${String((index % 7) + 1)}T08:00:00Z`
		})
	}
}

// Date filters on the drive's listing tools (the code resolves the period, the model only names it), and a
// note question the notes' own edited dates answer.
export const DATE_FILTER_CASES: EvalCase[] = [
	{
		id: "dates-search-last-year-en",
		category: "dates",
		lang: "en",
		prompt: "Find the invoices I got last year",
		routes: ["files", "insights"],
		calls: [
			{
				tool: "search_items",
				args: {
					query: {
						$includes: ["invoice"]
					},
					period: {
						$oneOf: ["last_year", "custom"]
					}
				}
			}
		],
		answerIncludes: [["telekom"], ["stadtwerke"], ["hetzner"]],
		answerExcludes: ["2025-02", "2023-12"],
		expectsText: true
	},
	{
		id: "dates-list-this-month-de",
		category: "dates",
		lang: "de",
		prompt: "Welche Dateien im Ordner Downloads habe ich diesen Monat geändert?",
		routes: ["files", "insights"],
		calls: [
			{
				tool: "list_directory",
				args: {
					period: {
						$oneOf: ["this_month", "custom"]
					}
				}
			}
		],
		answerIncludes: [["report_q3"], ["報告書"]],
		answerExcludes: ["setup_installer", "tugas_kuliah", "receipt_amazon", "отчёт"],
		expectsText: true
	},
	{
		id: "dates-list-this-year-ru",
		category: "dates",
		lang: "ru",
		prompt: "Какие файлы в папке Documents я изменил в этом году?",
		routes: ["files", "insights"],
		calls: [
			{
				tool: "list_directory",
				args: {
					period: {
						$oneOf: ["this_year", "custom"]
					}
				}
			}
		],
		answerIncludes: [["cv_2025"], ["tax_notes"], ["steuererklarung"], ["旅行计划"]],
		answerExcludes: ["old_draft"],
		expectsText: true
	},
	{
		id: "dates-notes-last-week-es",
		category: "dates",
		lang: "es",
		prompt: "¿Qué notas edité la semana pasada?",
		setup: state =>
			setEdited(state, {
				n4: "2025-10-01T18:00:00Z",
				n10: "2025-09-29T09:00:00Z"
			}),
		routes: ["notes"],
		calls: [
			{
				tool: "list_notes"
			}
		],
		answerIncludes: [["tortilla"], ["ideias"]],
		answerExcludes: ["groceries", "einkaufsliste", "meeting"],
		expectsText: true
	},
	{
		id: "dates-recent-photos-ja",
		category: "dates",
		lang: "ja",
		prompt: "今週アップロードした写真は？",
		routes: ["insights"],
		calls: [
			{
				tool: "recent_items"
			}
		],
		answerIncludes: [["oma"]],
		answerExcludes: ["img_2041", "beach_sunset"],
		expectsText: true
	},
	{
		id: "dates-uploaded-pdfs-tr",
		category: "dates",
		lang: "tr",
		prompt: "Geçen ay yüklediğim PDF dosyaları hangileri?",
		routes: ["insights", "files"],
		calls: [
			{
				anyOf: [
					{
						tool: "recent_items"
					},
					{
						tool: "search_items"
					}
				]
			}
		],
		answerIncludes: [["mietvertrag"], ["فاتورة"], ["किराया"]],
		answerExcludes: ["report_q3", "報告書", "invoice"],
		expectsText: true
	},
	{
		id: "dates-custom-month-fr",
		category: "dates",
		lang: "fr",
		prompt: "Quels fichiers du dossier Documents ai-je modifiés en mai 2025 ?",
		routes: ["insights", "files"],
		calls: [
			{
				tool: "list_directory",
				args: {
					period: "custom",
					from: {
						$includes: ["2025-05"]
					}
				}
			}
		],
		answerIncludes: [["cv_2025"]],
		answerExcludes: ["steuererklarung", "tax_notes", "旅行计划"],
		expectsText: true
	},
	{
		id: "dates-custom-range-en",
		category: "dates",
		lang: "en",
		prompt: "Show me the invoices dated between 1 January and 30 June 2024",
		routes: ["insights", "files"],
		calls: [
			{
				tool: "search_items",
				args: {
					query: {
						$includes: ["invoice"]
					},
					period: "custom",
					from: {
						$includes: ["2024-01"]
					},
					to: {
						$includes: ["2024-06"]
					}
				}
			}
		],
		answerIncludes: [["telekom"], ["stadtwerke"]],
		answerExcludes: ["hetzner", "2025-02"],
		expectsText: true
	},
	{
		id: "dates-notes-many-last-month-en",
		category: "dates",
		lang: "en",
		prompt: "Which notes did I edit last month?",
		setup: manyNotes,
		routes: ["notes"],
		calls: [
			{
				tool: "list_notes",
				args: {
					period: {
						$oneOf: ["last_month", "custom"]
					}
				}
			}
		],
		answerIncludes: [["tortilla"], ["ideias"]],
		answerExcludes: ["daily log", "groceries"],
		expectsText: true
	},
	{
		id: "dates-notes-many-last-month-pt",
		category: "dates",
		lang: "pt",
		prompt: "Quais notas eu editei no mês passado?",
		setup: manyNotes,
		routes: ["notes"],
		calls: [
			{
				tool: "list_notes",
				args: {
					period: {
						$oneOf: ["last_month", "custom"]
					}
				}
			}
		],
		answerIncludes: [["tortilla"], ["ideias"]],
		answerExcludes: ["daily log"],
		expectsText: true
	},
	{
		id: "dates-photos-taken-fr",
		category: "dates",
		lang: "fr",
		prompt: "Quelles photos ai-je prises en mai 2024 ?",
		routes: ["insights"],
		calls: [
			{
				tool: "list_photos"
			}
		],
		answerIncludes: [["img_2041", "img_2042", "beach_sunset"]],
		answerExcludes: ["oma", "weihnachten"],
		expectsText: true
	},
	{
		id: "dates-photos-december-zh",
		category: "dates",
		lang: "zh",
		prompt: "去年12月我拍了哪些照片或视频？",
		routes: ["insights"],
		calls: [
			{
				tool: "list_photos"
			}
		],
		answerIncludes: [["weihnachten"]],
		answerExcludes: ["oma", "img_2041", "beach_sunset"],
		expectsText: true
	}
]

export const INSIGHTS_CASES: EvalCase[] = [
	{
		id: "insights-largest-dir-en",
		category: "insights",
		lang: "en",
		prompt: "Which of my folders takes up the most space?",
		routes: ["insights", "files"],
		calls: [
			{
				anyOf: [
					{
						tool: "list_directory",
						args: {
							sort: "largest"
						}
					},
					{
						tool: "get_item_info",
						args: {
							handle: "d4"
						}
					}
				]
			}
		],
		answerIncludes: [["photos"]],
		expectsText: true
	},
	{
		id: "insights-largest-dir-ja",
		category: "insights",
		lang: "ja",
		prompt: "一番容量が大きいフォルダはどれ？",
		routes: ["insights", "files"],
		calls: [
			{
				anyOf: [
					{
						tool: "list_directory",
						args: {
							sort: "largest"
						}
					},
					{
						tool: "get_item_info",
						args: {
							handle: "d4"
						}
					}
				]
			}
		],
		answerIncludes: [["photos"]],
		expectsText: true
	},
	{
		id: "insights-most-space-documents-de",
		category: "insights",
		lang: "de",
		prompt: "Was belegt im Ordner Documents am meisten Platz?",
		routes: ["insights", "files"],
		calls: [
			{
				tool: "list_directory",
				args: {
					sort: "largest"
				}
			}
		],
		answerIncludes: [["vertrage"]],
		expectsText: true
	},
	{
		id: "insights-largest-downloads-en",
		category: "insights",
		lang: "en",
		prompt: "What are the biggest files in my Downloads folder?",
		routes: ["insights", "files"],
		calls: [
			{
				tool: "list_directory",
				args: {
					sort: "largest"
				}
			}
		],
		answerIncludes: [["setup_installer"]],
		expectsText: true
	},
	{
		id: "insights-largest-music-fr",
		category: "insights",
		lang: "fr",
		prompt: "Quel est le plus gros fichier de ma musique ?",
		routes: ["insights", "files"],
		calls: [
			{
				tool: "list_directory",
				args: {
					sort: "largest"
				}
			}
		],
		answerIncludes: [["clair de lune"]],
		expectsText: true
	},
	{
		id: "insights-oldest-pt",
		category: "insights",
		lang: "pt",
		prompt: "Qual é o arquivo mais antigo na pasta Escola?",
		routes: ["insights", "files"],
		calls: [
			{
				tool: "list_directory",
				args: {
					sort: "oldest"
				}
			}
		],
		answerIncludes: [["trabalho_de_historia"]],
		expectsText: true
	},
	{
		id: "insights-videos-tr",
		category: "insights",
		lang: "tr",
		prompt: "Familie klasöründe hangi videolar var?",
		routes: ["insights", "files"],
		calls: [
			{
				tool: "list_directory"
			}
		],
		answerIncludes: [["weihnachten"]],
		answerExcludes: ["oma_geburtstag"],
		expectsText: true
	},
	{
		id: "insights-dir-size-es",
		category: "insights",
		lang: "es",
		prompt: "¿Cuánto ocupa la carpeta Photos?",
		routes: ["insights", "files"],
		calls: [
			{
				tool: "get_item_info",
				args: {
					handle: "d4"
				}
			}
		],
		answerIncludes: [["195"]],
		expectsText: true
	},
	{
		id: "insights-dir-count-zh",
		category: "insights",
		lang: "zh",
		prompt: "Downloads 文件夹里有多少个文件？",
		routes: ["insights", "files"],
		calls: [
			{
				anyOf: [
					{
						tool: "get_item_info",
						args: {
							handle: "d13"
						}
					},
					{
						tool: "list_directory"
					}
				]
			}
		],
		answerIncludes: [["8", "八"]],
		expectsText: true
	},
	{
		id: "insights-storage-free-de",
		category: "insights",
		lang: "de",
		prompt: "Wie viel Speicherplatz habe ich noch frei?",
		routes: ["insights"],
		calls: [
			{
				tool: "get_storage"
			}
		],
		answerIncludes: [["6.8", "6,8"]],
		expectsText: true
	},
	{
		id: "insights-storage-used-ru",
		category: "insights",
		lang: "ru",
		prompt: "Сколько места в облаке я уже использовал?",
		routes: ["insights"],
		calls: [
			{
				tool: "get_storage"
			}
		],
		answerIncludes: [["3.2", "3,2", "32"]],
		expectsText: true
	},
	{
		id: "insights-storage-percent-hi",
		category: "insights",
		lang: "hi",
		prompt: "मेरा स्टोरेज कितने प्रतिशत भरा है?",
		routes: ["insights"],
		calls: [
			{
				tool: "get_storage"
			}
		],
		answerIncludes: [["32"]],
		expectsText: true
	},
	{
		id: "insights-storage-plan-en",
		category: "insights",
		lang: "en",
		prompt: "Am I running out of space?",
		routes: ["insights"],
		calls: [
			{
				tool: "get_storage"
			}
		],
		answerIncludes: [["6.8", "32", "3.2"]],
		expectsText: true
	},
	{
		id: "insights-recent-ja",
		category: "insights",
		lang: "ja",
		prompt: "最近アップロードしたファイルは何？",
		routes: ["insights"],
		calls: [
			{
				tool: "recent_items"
			}
		],
		answerIncludes: [["oma", "report_q3", "報告書"]],
		expectsText: true
	},
	{
		id: "insights-last-upload-ar",
		category: "insights",
		lang: "ar",
		prompt: "ما آخر ملف قمت برفعه؟",
		routes: ["insights"],
		calls: [
			{
				tool: "recent_items"
			}
		],
		answerIncludes: [["oma"]],
		expectsText: true
	},
	{
		id: "insights-largest-upload-id",
		category: "insights",
		lang: "id",
		prompt: "File apa yang paling besar yang saya unggah bulan ini?",
		routes: ["insights"],
		calls: [
			{
				tool: "recent_items"
			}
		],
		answerIncludes: [["oma"]],
		expectsText: true
	},
	{
		id: "insights-largest-files-en",
		category: "insights",
		lang: "en",
		prompt: "What are my largest files?",
		routes: ["insights"],
		// Graded on the answer alone: the largest file sits two directories deep, so neither Recents
		// (setup_installer.dmg) nor the root listing (its directories, README.txt) names it.
		answerIncludes: [["weihnachten"]],
		answerExcludes: ["readme"],
		expectsText: true
	},
	{
		id: "insights-largest-files-de",
		category: "insights",
		lang: "de",
		prompt: "Was sind meine größten Dateien?",
		routes: ["insights"],
		// Graded on the answer alone: the largest file sits two directories deep, so neither Recents
		// (setup_installer.dmg) nor the root listing (its directories, README.txt) names it.
		answerIncludes: [["weihnachten"]],
		answerExcludes: ["readme"],
		expectsText: true
	},
	{
		id: "insights-largest-files-zh",
		category: "insights",
		lang: "zh",
		prompt: "我最大的文件是哪些？",
		routes: ["insights"],
		// Graded on the answer alone: the largest file sits two directories deep, so neither Recents
		// (setup_installer.dmg) nor the root listing (its directories, README.txt) names it.
		answerIncludes: [["weihnachten"]],
		answerExcludes: ["readme"],
		expectsText: true
	},
	{
		id: "insights-largest-files-ar",
		category: "insights",
		lang: "ar",
		prompt: "ما هي أكبر ملفاتي؟",
		routes: ["insights"],
		// Graded on the answer alone: the largest file sits two directories deep, so neither Recents
		// (setup_installer.dmg) nor the root listing (its directories, README.txt) names it.
		answerIncludes: [["weihnachten"]],
		answerExcludes: ["readme"],
		expectsText: true
	},
	{
		id: "insights-worked-recently-en",
		category: "insights",
		lang: "en",
		prompt: "What did I work on recently?",
		routes: ["insights"],
		calls: [
			{
				tool: "recent_items"
			}
		],
		answerIncludes: [["oma", "report_q3", "報告書", "roadmap", "meeting_notes"]],
		expectsText: true
	},
	{
		id: "insights-worked-recently-es",
		category: "insights",
		lang: "es",
		prompt: "¿En qué he trabajado últimamente?",
		routes: ["insights"],
		calls: [
			{
				tool: "recent_items"
			}
		],
		answerIncludes: [["oma", "report_q3", "報告書", "roadmap", "meeting_notes"]],
		expectsText: true
	},
	{
		id: "insights-worked-recently-fr",
		category: "insights",
		lang: "fr",
		prompt: "Sur quoi ai-je travaillé récemment ?",
		routes: ["insights"],
		calls: [
			{
				tool: "recent_items"
			}
		],
		answerIncludes: [["oma", "report_q3", "報告書", "roadmap", "meeting_notes"]],
		expectsText: true
	},
	{
		id: "insights-worked-recently-ru",
		category: "insights",
		lang: "ru",
		prompt: "Над чем я недавно работал?",
		routes: ["insights"],
		calls: [
			{
				tool: "recent_items"
			}
		],
		answerIncludes: [["oma", "report_q3", "報告書", "roadmap", "meeting_notes"]],
		expectsText: true
	},
	{
		id: "insights-worked-recently-tr",
		category: "insights",
		lang: "tr",
		prompt: "Son zamanlarda hangi dosyalar üzerinde çalıştım?",
		routes: ["insights"],
		calls: [
			{
				tool: "recent_items"
			}
		],
		answerIncludes: [["oma", "report_q3", "報告書", "roadmap", "meeting_notes"]],
		expectsText: true
	}
]
