import type { EvalCase } from "./types"

// Lookups followed by actions; graded on the end state of the fake app.
export const MULTI_STEP_CASES: EvalCase[] = [
	{
		id: "multi-move-invoices-2024-en",
		category: "multi_step",
		lang: "en",
		prompt: "Move the invoices from 2024 into Archive",
		routes: ["organize"],
		state: [
			{
				item: "f1",
				parent: "d8"
			},
			{
				item: "f2",
				parent: "d8"
			},
			{
				item: "f3",
				parent: "d8"
			}
		],
		mayChange: ["f5"],
		note: "f5 is the German-named 2024 invoice; moving it is fine, f4 (2025) must stay"
	},
	{
		id: "multi-move-telekom-de",
		category: "multi_step",
		lang: "de",
		prompt: "Verschiebe alle Telekom-Rechnungen ins Archiv",
		routes: ["organize"],
		state: [
			{
				item: "f1",
				parent: "d8"
			},
			{
				item: "f4",
				parent: "d8"
			}
		],
		mayChange: ["f26"]
	},
	{
		id: "multi-create-move-es",
		category: "multi_step",
		lang: "es",
		prompt: "Crea una carpeta llamada Viajes dentro de Projects y mueve ahí la carpeta Proyecto Barcelona",
		routes: ["organize"],
		state: [
			{
				newDirectory: {
					parent: "d10",
					name: "Viajes"
				}
			},
			{
				item: "d12",
				parent: "new:Viajes"
			}
		]
	},
	{
		id: "multi-rename-fr",
		category: "multi_step",
		lang: "fr",
		prompt: "Renomme le fichier CV_2025.docx en CV_Sam_2025.docx",
		routes: ["organize"],
		state: [
			{
				item: "f8",
				name: "CV_Sam_2025.docx"
			}
		]
	},
	{
		id: "multi-create-move-pt",
		category: "multi_step",
		lang: "pt",
		prompt: "Cria uma pasta chamada Faculdade na raiz do drive e move o trabalho de história pra lá",
		routes: ["organize"],
		state: [
			{
				newDirectory: {
					parent: "d0",
					name: "Faculdade"
				}
			},
			{
				item: "f38",
				parent: "new:Faculdade"
			}
		]
	},
	{
		id: "multi-move-ru",
		category: "multi_step",
		lang: "ru",
		prompt: "Перемести файл отчёт_2024.docx в папку Documents",
		routes: ["organize"],
		state: [
			{
				item: "f35",
				parent: "d1"
			}
		]
	},
	{
		id: "multi-move-tr",
		category: "multi_step",
		lang: "tr",
		prompt: "Downloads klasöründeki Amazon faturasını Invoices klasörüne taşı",
		routes: ["organize"],
		state: [
			{
				item: "f34",
				parent: "d2"
			}
		]
	},
	{
		id: "multi-move-ar",
		category: "multi_step",
		lang: "ar",
		prompt: "انقل فاتورة الكهرباء إلى مجلد Invoices",
		routes: ["organize"],
		state: [
			{
				item: "f42",
				parent: "d2"
			}
		]
	},
	{
		id: "multi-copy-hi",
		category: "multi_step",
		lang: "hi",
		prompt: "किराया रसीद वाली फ़ाइल को Documents फ़ोल्डर में कॉपी करो",
		routes: ["organize"],
		state: [
			{
				copyOf: "f43",
				parent: "d1"
			}
		]
	},
	{
		id: "multi-create-move-id",
		category: "multi_step",
		lang: "id",
		prompt: "Buat folder Kuliah di dalam Documents, lalu pindahkan tugas_kuliah.docx ke sana",
		routes: ["organize"],
		state: [
			{
				newDirectory: {
					parent: "d1",
					name: "Kuliah"
				}
			},
			{
				item: "f37",
				parent: "new:Kuliah"
			}
		]
	},
	{
		id: "multi-move-zh",
		category: "multi_step",
		lang: "zh",
		prompt: "把 report_Q3.pdf 移到 Documents 文件夹",
		routes: ["organize"],
		state: [
			{
				item: "f33",
				parent: "d1"
			}
		]
	},
	{
		id: "multi-copy-ja",
		category: "multi_step",
		lang: "ja",
		prompt: "報告書のPDFをArchiveフォルダにコピーして",
		routes: ["organize"],
		state: [
			{
				copyOf: "f36",
				parent: "d8"
			}
		]
	},
	{
		id: "multi-add-lofi-en",
		category: "multi_step",
		lang: "en",
		prompt: "Add all the lofi tracks to my Workout playlist",
		routes: ["audio"],
		state: [
			{
				playlist: "p1",
				includes: ["f22", "f23"]
			}
		]
	},
	{
		id: "multi-add-tracks-de",
		category: "multi_step",
		lang: "de",
		prompt: "Füge Kalinka und Bella Ciao zur Playlist Chill hinzu",
		routes: ["audio"],
		state: [
			{
				playlist: "p2",
				includes: ["f19", "f25"]
			}
		]
	},
	{
		id: "multi-note-flags-en",
		category: "multi_step",
		lang: "en",
		prompt: "Archive my note 'Meeting 2025-10-01' and unpin the Groceries note",
		routes: ["notes"],
		state: [
			{
				note: "n3",
				archived: true
			},
			{
				note: "n1",
				pinned: false
			}
		]
	},
	{
		id: "multi-favorite-open-note-es",
		category: "multi_step",
		lang: "es",
		prompt: "Marca como favorita la nota de la tortilla y luego ábrela",
		routes: ["notes"],
		state: [
			{
				note: "n4",
				favorite: true
			},
			{
				screen: "note",
				opened: "n4"
			}
		]
	},
	{
		id: "multi-unfavorite-fr",
		category: "multi_step",
		lang: "fr",
		prompt: "Retire le dossier Lisbon 2024 et le fichier CV_2025.docx des favoris",
		routes: ["files"],
		state: [
			{
				item: "d5",
				favorite: false
			},
			{
				item: "f8",
				favorite: false
			}
		]
	},
	{
		id: "multi-copy-rename-pt",
		category: "multi_step",
		lang: "pt",
		prompt: "Copia o arquivo redação_enem.txt para a pasta Documents e renomeia a cópia para redacao_final.txt",
		routes: ["organize"],
		state: [
			{
				copyOf: "f39",
				parent: "d1",
				name: "redacao_final.txt"
			}
		]
	},
	{
		id: "multi-reveal-en",
		category: "multi_step",
		lang: "en",
		prompt: "Find the Hetzner invoice and show it to me in the drive",
		routes: ["files"],
		calls: [
			{
				tool: "reveal_item",
				args: {
					handle: "f3"
				}
			}
		]
	},
	{
		id: "multi-create-move-ru",
		category: "multi_step",
		lang: "ru",
		prompt: "Создай в Documents папку «Отчёты» и перемести туда отчёт_2024.docx",
		routes: ["organize"],
		state: [
			{
				newDirectory: {
					parent: "d1",
					name: "Отчёты"
				}
			},
			{
				item: "f35",
				parent: "new:Отчёты"
			}
		]
	},
	{
		id: "multi-move-directory-tr",
		category: "multi_step",
		lang: "tr",
		prompt: "Tarifler klasörünü Documents klasörünün içine taşı",
		routes: ["organize"],
		state: [
			{
				item: "d15",
				parent: "d1"
			}
		]
	},
	{
		id: "multi-favorite-photos-ja",
		category: "multi_step",
		lang: "ja",
		prompt: "Lisbon 2024フォルダの写真を全部お気に入りに追加して",
		routes: ["files"],
		state: [
			{
				item: "f11",
				favorite: true
			},
			{
				item: "f12",
				favorite: true
			},
			{
				item: "f13",
				favorite: true
			}
		]
	},
	{
		id: "multi-draft-en",
		category: "multi_step",
		lang: "en",
		prompt: "Write Anna a message that I paid the Hetzner invoice",
		routes: ["chats"],
		state: [
			{
				draft: "c1",
				textIncludes: ["hetzner"]
			}
		]
	},
	{
		id: "multi-draft-de",
		category: "multi_step",
		lang: "de",
		prompt: "Schreib Carlos, dass ich ihm das Budget heute schicke",
		routes: ["chats"],
		state: [
			{
				draft: "c3",
				textIncludes: ["budget", "presupuesto"]
			}
		]
	},
	{
		id: "multi-play-shuffle-zh",
		category: "multi_step",
		lang: "zh",
		prompt: "播放 Road trip 歌单，然后打开随机播放",
		routes: ["audio"],
		state: [
			{
				playing: "f16",
				shuffle: true
			}
		]
	},
	{
		id: "multi-move-photos-id",
		category: "multi_step",
		lang: "id",
		prompt: "Pindahkan semua foto dari folder Lisbon 2024 ke folder Familie",
		routes: ["organize"],
		state: [
			{
				item: "f11",
				parent: "d6"
			},
			{
				item: "f12",
				parent: "d6"
			},
			{
				item: "f13",
				parent: "d6"
			}
		]
	},
	{
		id: "multi-create-directory-hi",
		category: "multi_step",
		lang: "hi",
		prompt: "Archive फ़ोल्डर के अंदर 2025 नाम का नया फ़ोल्डर बनाओ",
		routes: ["organize"],
		state: [
			{
				newDirectory: {
					parent: "d8",
					name: "2025"
				}
			}
		]
	},
	{
		id: "multi-move-ar-2",
		category: "multi_step",
		lang: "ar",
		prompt: "انقل ملف Steuererklärung_2024.pdf إلى مجلد Verträge",
		routes: ["organize"],
		state: [
			{
				item: "f10",
				parent: "d3"
			}
		]
	}
]
