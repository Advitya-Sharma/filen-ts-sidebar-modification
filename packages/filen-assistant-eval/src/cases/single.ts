import { inDirectory, PLAYING, selected } from "./contexts"
import type { EvalCase } from "./types"

// One call does the job; the arguments come from the prompt or the context.
export const SIMPLE_CASES: EvalCase[] = [
	{
		id: "simple-pause-en",
		category: "simple",
		lang: "en",
		prompt: "Pause the music",
		context: PLAYING,
		routes: ["audio"],
		calls: [
			{
				tool: "audio_control",
				args: {
					action: "pause"
				}
			}
		]
	},
	{
		id: "simple-next-de",
		category: "simple",
		lang: "de",
		prompt: "Spiel das nächste Lied",
		context: PLAYING,
		routes: ["audio"],
		calls: [
			{
				tool: "audio_control",
				args: {
					action: "next"
				}
			}
		]
	},
	{
		id: "simple-shuffle-es",
		category: "simple",
		lang: "es",
		prompt: "Activa el modo aleatorio",
		context: PLAYING,
		routes: ["audio"],
		calls: [
			{
				tool: "audio_control",
				args: {
					action: "shuffle_on"
				}
			}
		]
	},
	{
		id: "simple-settings-fr",
		category: "simple",
		lang: "fr",
		prompt: "Ouvre les paramètres",
		routes: ["navigation"],
		calls: [
			{
				tool: "open_route",
				args: {
					route: "settings"
				}
			}
		]
	},
	{
		id: "simple-trash-route-pt",
		category: "simple",
		lang: "pt",
		prompt: "Abre a lixeira pra mim",
		routes: ["navigation"],
		calls: [
			{
				tool: "open_route",
				args: {
					route: "trash"
				}
			}
		]
	},
	{
		id: "simple-search-ru",
		category: "simple",
		lang: "ru",
		prompt: "Найди мой отчёт за 2024 год",
		routes: ["files"],
		calls: [
			{
				tool: "search_items",
				args: {
					query: {
						$includes: ["отчет", "2024", "report"]
					}
				}
			}
		]
	},
	{
		id: "simple-notes-route-tr",
		category: "simple",
		lang: "tr",
		prompt: "Notlar bölümünü aç",
		routes: ["navigation"],
		calls: [
			{
				tool: "open_route",
				args: {
					route: "notes"
				}
			}
		]
	},
	{
		id: "simple-search-ar",
		category: "simple",
		lang: "ar",
		prompt: "ابحث عن فاتورة الكهرباء",
		routes: ["files"],
		calls: [
			{
				tool: "search_items",
				args: {
					query: {
						$includes: ["فاتورة", "الكهرباء", "electric"]
					}
				}
			}
		]
	},
	{
		id: "simple-playlists-hi",
		category: "simple",
		lang: "hi",
		prompt: "मेरी सारी प्लेलिस्ट दिखाओ",
		routes: ["audio", "navigation"],
		calls: [
			{
				anyOf: [
					{
						tool: "list_playlists"
					},
					{
						tool: "open_route",
						args: {
							route: "audio"
						}
					}
				]
			}
		]
	},
	{
		id: "simple-search-id",
		category: "simple",
		lang: "id",
		prompt: "Cari file tugas kuliah saya",
		routes: ["files"],
		calls: [
			{
				tool: "search_items",
				args: {
					query: {
						$includes: ["tugas", "kuliah"]
					}
				}
			}
		]
	},
	{
		id: "simple-chats-route-zh",
		category: "simple",
		lang: "zh",
		prompt: "打开聊天页面",
		routes: ["navigation", "chats"],
		calls: [
			{
				anyOf: [
					{
						tool: "open_route",
						args: {
							route: "chats"
						}
					},
					{
						tool: "list_chats"
					}
				]
			}
		]
	},
	{
		id: "simple-pause-ja",
		category: "simple",
		lang: "ja",
		prompt: "音楽を一時停止して",
		context: PLAYING,
		routes: ["audio"],
		calls: [
			{
				tool: "audio_control",
				args: {
					action: "pause"
				}
			}
		]
	},
	{
		id: "simple-create-note-en",
		category: "simple",
		lang: "en",
		prompt: "Create a note called Packing list",
		routes: ["notes"],
		calls: [
			{
				tool: "create_note",
				args: {
					title: "Packing list"
				}
			}
		],
		state: [
			{
				newNote: {
					titleIncludes: ["packing list"]
				}
			}
		]
	},
	{
		id: "simple-create-directory-de",
		category: "simple",
		lang: "de",
		prompt: "Erstelle hier einen neuen Ordner „Steuern 2025“",
		context: inDirectory("d1"),
		routes: ["organize"],
		calls: [
			{
				tool: "create_directory",
				args: {
					parent: "d1",
					name: "Steuern 2025"
				}
			}
		],
		state: [
			{
				newDirectory: {
					parent: "d1",
					name: "Steuern 2025"
				}
			}
		]
	},
	{
		id: "simple-rename-es",
		category: "simple",
		lang: "es",
		prompt: "Renombra este archivo a factura_hetzner.pdf",
		context: selected({
			directory: "d2",
			items: ["f3"]
		}),
		routes: ["organize"],
		calls: [
			{
				tool: "rename_item",
				args: {
					handle: "f3",
					new_name: "factura_hetzner.pdf"
				}
			}
		],
		state: [
			{
				item: "f3",
				name: "factura_hetzner.pdf"
			}
		]
	},
	{
		id: "simple-favorite-fr",
		category: "simple",
		lang: "fr",
		prompt: "Ajoute ce fichier aux favoris",
		context: selected({
			directory: "d13",
			items: ["f33"]
		}),
		routes: ["files"],
		calls: [
			{
				tool: "set_favorite",
				args: {
					handles: ["f33"],
					favorite: true
				}
			}
		],
		state: [
			{
				item: "f33",
				favorite: true
			}
		]
	},
	{
		id: "simple-info-pt",
		category: "simple",
		lang: "pt",
		prompt: "Mostra as informações deste arquivo",
		context: selected({
			directory: "d14",
			items: ["f38"]
		}),
		routes: ["files"],
		calls: [
			{
				tool: "get_item_info",
				args: {
					handle: "f38"
				}
			}
		],
		expectsText: true
	},
	{
		id: "simple-security-ru",
		category: "simple",
		lang: "ru",
		prompt: "Открой настройки безопасности",
		routes: ["navigation"],
		calls: [
			{
				tool: "open_route",
				args: {
					route: "settings_security"
				}
			}
		]
	},
	{
		id: "simple-read-tr",
		category: "simple",
		lang: "tr",
		prompt: "Bu dosyayı oku ve kısaca anlat",
		context: selected({
			directory: "d15",
			items: ["f40"]
		}),
		routes: ["files"],
		calls: [
			{
				tool: "read_text_file",
				args: {
					handle: "f40"
				}
			}
		],
		answerIncludes: [["mercimek"]],
		expectsText: true
	},
	{
		id: "simple-next-ar",
		category: "simple",
		lang: "ar",
		prompt: "شغّل الأغنية التالية",
		context: PLAYING,
		routes: ["audio"],
		calls: [
			{
				tool: "audio_control",
				args: {
					action: "next"
				}
			}
		]
	},
	{
		id: "simple-list-notes-hi",
		category: "simple",
		lang: "hi",
		prompt: "मेरे नोट्स में 'यात्रा' खोजो",
		routes: ["notes"],
		calls: [
			{
				tool: "list_notes",
				args: {
					query: {
						$includes: ["यात्रा"]
					}
				}
			}
		]
	},
	{
		id: "simple-appearance-id",
		category: "simple",
		lang: "id",
		prompt: "Buka pengaturan tampilan",
		routes: ["navigation"],
		calls: [
			{
				tool: "open_route",
				args: {
					route: "settings_appearance"
				}
			}
		]
	},
	{
		id: "simple-search-zh",
		category: "simple",
		lang: "zh",
		prompt: "搜索所有的发票",
		routes: ["files"],
		calls: [
			{
				tool: "search_items",
				args: {
					query: {
						$includes: ["发票", "invoice", "rechnung"]
					}
				}
			}
		]
	},
	{
		id: "simple-unfavorite-ja",
		category: "simple",
		lang: "ja",
		prompt: "このファイルをお気に入りから外して",
		context: selected({
			directory: "d1",
			items: ["f8"]
		}),
		routes: ["files"],
		calls: [
			{
				tool: "set_favorite",
				args: {
					handles: ["f8"],
					favorite: false
				}
			}
		],
		state: [
			{
				item: "f8",
				favorite: false
			}
		]
	},
	{
		id: "simple-list-downloads-en",
		category: "simple",
		lang: "en",
		prompt: "What's in my Downloads folder?",
		routes: ["files"],
		calls: [
			{
				tool: "list_directory",
				args: {
					path: {
						$includes: ["downloads"]
					}
				}
			}
		],
		answerIncludes: [["report_q3", "setup_installer", "receipt"]],
		expectsText: true
	},
	{
		id: "simple-stop-de",
		category: "simple",
		lang: "de",
		prompt: "Stopp die Wiedergabe",
		context: PLAYING,
		routes: ["audio"],
		calls: [
			{
				tool: "audio_control",
				args: {
					action: "stop"
				}
			}
		]
	},
	{
		id: "simple-repeat-en",
		category: "simple",
		lang: "en",
		prompt: "Turn on repeat",
		context: PLAYING,
		routes: ["audio"],
		calls: [
			{
				tool: "audio_control",
				args: {
					action: "repeat_on"
				}
			}
		]
	},
	{
		id: "simple-list-chats-es",
		category: "simple",
		lang: "es",
		prompt: "¿Con quién tengo chats abiertos?",
		routes: ["chats"],
		calls: [
			{
				tool: "list_chats"
			}
		],
		answerIncludes: [["anna", "carlos", "family"]],
		expectsText: true
	},
	{
		id: "simple-checklist-fr",
		category: "simple",
		lang: "fr",
		prompt: "Crée une checklist « Courses » avec pain, lait et beurre",
		routes: ["notes"],
		calls: [
			{
				tool: "create_note",
				args: {
					title: {
						$includes: ["courses"]
					},
					type: "checklist"
				}
			}
		],
		state: [
			{
				newNote: {
					titleIncludes: ["courses"],
					type: "checklist",
					contentIncludes: ["pain", "lait", "beurre"]
				}
			}
		]
	}
]

// Several tools fit the group; the model has to pick the right one, usually after one lookup.
export const MULTIPLE_CASES: EvalCase[] = [
	{
		id: "multiple-play-playlist-en",
		category: "multiple",
		lang: "en",
		prompt: "Play my Workout playlist",
		routes: ["audio"],
		calls: [
			{
				tool: "play",
				args: {
					handle: "p1"
				}
			}
		],
		state: [
			{
				playing: "f18"
			}
		]
	},
	{
		id: "multiple-open-playlist-de",
		category: "multiple",
		lang: "de",
		prompt: "Öffne die Playlist Road trip",
		routes: ["audio"],
		calls: [
			{
				tool: "open_playlist",
				args: {
					handle: "p3"
				}
			}
		],
		forbidden: ["play"]
	},
	{
		id: "multiple-open-note-es",
		category: "multiple",
		lang: "es",
		prompt: "Abre mi nota de la tortilla",
		routes: ["notes"],
		calls: [
			{
				tool: "open_note",
				args: {
					handle: "n4"
				}
			}
		]
	},
	{
		id: "multiple-reveal-fr",
		category: "multiple",
		lang: "fr",
		prompt: "Montre-moi où se trouve le fichier itinerario.txt",
		routes: ["files"],
		calls: [
			{
				tool: "reveal_item",
				args: {
					handle: "f31"
				}
			}
		]
	},
	{
		id: "multiple-pin-note-pt",
		category: "multiple",
		lang: "pt",
		prompt: "Fixa a nota Ideias de app no topo",
		routes: ["notes"],
		calls: [
			{
				tool: "set_note_flags",
				args: {
					handle: "n10",
					pinned: true
				}
			}
		],
		state: [
			{
				note: "n10",
				pinned: true
			}
		]
	},
	{
		id: "multiple-open-chat-ru",
		category: "multiple",
		lang: "ru",
		prompt: "Открой чат с Марией",
		routes: ["chats"],
		calls: [
			{
				tool: "open_chat",
				args: {
					handle: "c6"
				}
			}
		]
	},
	{
		id: "multiple-open-chat-tr",
		category: "multiple",
		lang: "tr",
		prompt: "Ahmet ile olan sohbeti aç",
		routes: ["chats"],
		calls: [
			{
				tool: "open_chat",
				args: {
					handle: "c7"
				}
			}
		]
	},
	{
		id: "multiple-open-note-ar",
		category: "multiple",
		lang: "ar",
		prompt: "افتح ملاحظة خطة السفر",
		routes: ["notes"],
		calls: [
			{
				tool: "open_note",
				args: {
					handle: "n13"
				}
			}
		]
	},
	{
		id: "multiple-play-track-hi",
		category: "multiple",
		lang: "hi",
		prompt: "Tum Hi Ho गाना चलाओ",
		routes: ["audio"],
		calls: [
			{
				tool: "play",
				args: {
					handle: "f20"
				}
			}
		],
		state: [
			{
				playing: "f20"
			}
		]
	},
	{
		id: "multiple-archive-note-id",
		category: "multiple",
		lang: "id",
		prompt: "Arsipkan catatan rapat saya",
		routes: ["notes"],
		calls: [
			{
				tool: "set_note_flags",
				args: {
					handle: "n11",
					archived: true
				}
			}
		],
		state: [
			{
				note: "n11",
				archived: true
			}
		]
	},
	{
		id: "multiple-play-playlist-zh",
		category: "multiple",
		lang: "zh",
		prompt: "播放 Chill 歌单",
		routes: ["audio"],
		calls: [
			{
				tool: "play",
				args: {
					handle: "p2"
				}
			}
		],
		state: [
			{
				playing: "f17"
			}
		]
	},
	{
		id: "multiple-open-chat-ja",
		category: "multiple",
		lang: "ja",
		prompt: "Yukiさんとのチャットを開いて",
		routes: ["chats"],
		calls: [
			{
				tool: "open_chat",
				args: {
					handle: "c5"
				}
			}
		]
	},
	{
		id: "multiple-add-track-en",
		category: "multiple",
		lang: "en",
		prompt: "Add Bella Ciao to my Road trip playlist",
		routes: ["audio"],
		calls: [
			{
				tool: "add_tracks_to_playlist",
				args: {
					playlist: "p3",
					handles: ["f25"]
				}
			}
		],
		state: [
			{
				playlist: "p3",
				includes: ["f25"]
			}
		]
	},
	{
		id: "multiple-favorite-note-de",
		category: "multiple",
		lang: "de",
		prompt: "Markiere die Notiz Einkaufsliste als Favorit",
		routes: ["notes"],
		calls: [
			{
				tool: "set_note_flags",
				args: {
					handle: "n2",
					favorite: true
				}
			}
		],
		state: [
			{
				note: "n2",
				favorite: true
			}
		]
	},
	{
		id: "multiple-copy-es",
		category: "multiple",
		lang: "es",
		prompt: "Copia presupuesto.xlsx a la carpeta Documents",
		routes: ["organize"],
		calls: [
			{
				tool: "copy_items",
				args: {
					handles: ["f30"],
					destination: "d1"
				}
			}
		],
		state: [
			{
				copyOf: "f30",
				parent: "d1"
			}
		]
	},
	{
		id: "multiple-read-file-fr",
		category: "multiple",
		lang: "fr",
		prompt: "Lis le fichier roadmap.md et dis-moi qui s'occupe de la page des tarifs",
		routes: ["files"],
		calls: [
			{
				tool: "read_text_file",
				args: {
					handle: "f27"
				}
			}
		],
		answerIncludes: [["lena"]],
		expectsText: true
	},
	{
		id: "multiple-play-track-pt",
		category: "multiple",
		lang: "pt",
		prompt: "Toca Garota de Ipanema",
		routes: ["audio"],
		calls: [
			{
				tool: "play",
				args: {
					handle: "f24"
				}
			}
		],
		state: [
			{
				playing: "f24"
			}
		]
	},
	{
		id: "multiple-read-note-en",
		category: "multiple",
		lang: "en",
		prompt: "What did we decide in the meeting on October 1st? It's in my notes.",
		routes: ["notes"],
		calls: [
			{
				tool: "read_note",
				args: {
					handle: "n3"
				}
			}
		],
		answerIncludes: [["15", "november"]],
		expectsText: true
	},
	{
		id: "multiple-reveal-ru",
		category: "multiple",
		lang: "ru",
		prompt: "Покажи, где лежит файл отчёт_2024.docx",
		routes: ["files"],
		calls: [
			{
				tool: "reveal_item",
				args: {
					handle: "f35"
				}
			}
		]
	},
	{
		id: "multiple-favorites-route-en",
		category: "multiple",
		lang: "en",
		prompt: "Take me to my favorites",
		routes: ["navigation"],
		calls: [
			{
				tool: "open_route",
				args: {
					route: "favorites"
				}
			}
		]
	},
	{
		id: "multiple-read-file-tr",
		category: "multiple",
		lang: "tr",
		prompt: "Mercimek çorbası tarifinde hangi malzemeler var?",
		routes: ["files", "notes"],
		calls: [
			{
				tool: "read_text_file",
				args: {
					handle: "f40"
				}
			}
		],
		answerIncludes: [["mercimek"], ["sogan", "havuc"]],
		expectsText: true
	},
	{
		id: "multiple-open-chat-en",
		category: "multiple",
		lang: "en",
		prompt: "Open the chat with the website team",
		routes: ["chats"],
		calls: [
			{
				tool: "open_chat",
				args: {
					handle: "c4"
				}
			}
		]
	}
]

// More than one call in the same turn, ideally emitted together.
export const PARALLEL_CASES: EvalCase[] = [
	{
		id: "parallel-pause-shuffle-en",
		category: "parallel",
		lang: "en",
		prompt: "Pause the music and turn on shuffle",
		context: PLAYING,
		routes: ["audio"],
		calls: [
			{
				tool: "audio_control",
				args: {
					action: "pause"
				}
			},
			{
				tool: "audio_control",
				args: {
					action: "shuffle_on"
				}
			}
		],
		state: [
			{
				playing: "f17",
				paused: true,
				shuffle: true
			}
		]
	},
	{
		id: "parallel-search-de",
		category: "parallel",
		lang: "de",
		prompt: "Suche nach Telekom und nach Vodafone",
		routes: ["files"],
		calls: [
			{
				tool: "search_items",
				args: {
					query: {
						$includes: ["telekom"]
					}
				}
			},
			{
				tool: "search_items",
				args: {
					query: {
						$includes: ["vodafone"]
					}
				}
			}
		]
	},
	{
		id: "parallel-search-es",
		category: "parallel",
		lang: "es",
		prompt: "Busca mis facturas y también mis contratos",
		routes: ["files"],
		calls: [
			{
				tool: "search_items",
				args: {
					query: {
						$includes: ["factura", "invoice", "rechnung"]
					}
				}
			},
			{
				tool: "search_items",
				args: {
					query: {
						$includes: ["contrat", "vertrag", "contract"]
					}
				}
			}
		]
	},
	{
		id: "parallel-pin-open-fr",
		category: "parallel",
		lang: "fr",
		prompt: "Épingle la note Idées cadeaux Noël et ouvre-la",
		routes: ["notes"],
		calls: [
			{
				tool: "set_note_flags",
				args: {
					handle: "n5",
					pinned: true
				}
			},
			{
				tool: "open_note",
				args: {
					handle: "n5"
				}
			}
		],
		state: [
			{
				note: "n5",
				pinned: true
			},
			{
				screen: "note",
				opened: "n5"
			}
		]
	},
	{
		id: "parallel-list-dirs-pt",
		category: "parallel",
		lang: "pt",
		prompt: "Lista o que tem nas pastas Escola e Downloads",
		routes: ["files"],
		calls: [
			{
				tool: "list_directory",
				args: {
					path: {
						$includes: ["escola"]
					}
				}
			},
			{
				tool: "list_directory",
				args: {
					path: {
						$includes: ["downloads"]
					}
				}
			}
		],
		answerIncludes: [["trabalho_de_historia", "redacao_enem"]],
		expectsText: true
	},
	{
		id: "parallel-create-notes-ru",
		category: "parallel",
		lang: "ru",
		prompt: "Создай две заметки: «Покупки» и «Дела на неделю»",
		routes: ["notes"],
		calls: [
			{
				tool: "create_note",
				args: {
					title: {
						$includes: ["покупки"]
					}
				}
			},
			{
				tool: "create_note",
				args: {
					title: {
						$includes: ["дела"]
					}
				}
			}
		],
		state: [
			{
				newNote: {
					titleIncludes: ["покупки"]
				}
			},
			{
				newNote: {
					titleIncludes: ["дела на неделю"]
				}
			}
		]
	},
	{
		id: "parallel-search-tr",
		category: "parallel",
		lang: "tr",
		prompt: "Kalinka ve Tum Hi Ho şarkılarını bul",
		routes: ["audio", "files"],
		calls: [
			{
				tool: "search_items",
				args: {
					query: {
						$includes: ["kalinka"]
					}
				}
			},
			{
				tool: "search_items",
				args: {
					query: {
						$includes: ["tum hi ho"]
					}
				}
			}
		]
	},
	{
		id: "parallel-create-directories-ar",
		category: "parallel",
		lang: "ar",
		prompt: "أنشئ مجلدين هنا: صور وفواتير",
		context: inDirectory("d1"),
		routes: ["organize"],
		calls: [
			{
				tool: "create_directory",
				args: {
					parent: "d1",
					name: "صور"
				}
			},
			{
				tool: "create_directory",
				args: {
					parent: "d1",
					name: "فواتير"
				}
			}
		],
		state: [
			{
				newDirectory: {
					parent: "d1",
					name: "صور"
				}
			},
			{
				newDirectory: {
					parent: "d1",
					name: "فواتير"
				}
			}
		]
	},
	{
		id: "parallel-search-hi",
		category: "parallel",
		lang: "hi",
		prompt: "Despacito और Kalinka दोनों गाने खोजो",
		routes: ["audio", "files"],
		calls: [
			{
				tool: "search_items",
				args: {
					query: {
						$includes: ["despacito"]
					}
				}
			},
			{
				tool: "search_items",
				args: {
					query: {
						$includes: ["kalinka"]
					}
				}
			}
		]
	},
	{
		id: "parallel-create-directories-id",
		category: "parallel",
		lang: "id",
		prompt: "Buat dua folder di sini: Kuliah dan Kerja",
		context: inDirectory("d13"),
		routes: ["organize"],
		calls: [
			{
				tool: "create_directory",
				args: {
					parent: "d13",
					name: "Kuliah"
				}
			},
			{
				tool: "create_directory",
				args: {
					parent: "d13",
					name: "Kerja"
				}
			}
		],
		state: [
			{
				newDirectory: {
					parent: "d13",
					name: "Kuliah"
				}
			},
			{
				newDirectory: {
					parent: "d13",
					name: "Kerja"
				}
			}
		]
	},
	{
		id: "parallel-search-directories-zh",
		category: "parallel",
		lang: "zh",
		prompt: "帮我找一下 Lisbon 2024 和 Familie 这两个文件夹",
		routes: ["files"],
		calls: [
			{
				tool: "search_items",
				args: {
					query: {
						$includes: ["lisbon"]
					}
				}
			},
			{
				tool: "search_items",
				args: {
					query: {
						$includes: ["familie"]
					}
				}
			}
		]
	},
	{
		id: "parallel-shuffle-repeat-ja",
		category: "parallel",
		lang: "ja",
		prompt: "シャッフルとリピートをオンにして",
		context: PLAYING,
		routes: ["audio"],
		calls: [
			{
				tool: "audio_control",
				args: {
					action: "shuffle_on"
				}
			},
			{
				tool: "audio_control",
				args: {
					action: "repeat_on"
				}
			}
		]
	},
	{
		id: "parallel-read-files-en",
		category: "parallel",
		lang: "en",
		prompt: "Read roadmap.md and the meeting notes file in the Website relaunch directory and tell me the launch date",
		routes: ["files"],
		calls: [
			{
				tool: "read_text_file",
				args: {
					handle: "f27"
				}
			},
			{
				tool: "read_text_file",
				args: {
					handle: "f28"
				}
			}
		],
		answerIncludes: [["15", "november"]],
		expectsText: true
	},
	{
		id: "parallel-note-flags-de",
		category: "parallel",
		lang: "de",
		prompt: "Hefte die Einkaufsliste an und archiviere die Notiz Meeting 2025-10-01",
		routes: ["notes"],
		calls: [
			{
				tool: "set_note_flags",
				args: {
					handle: "n2",
					pinned: true
				}
			},
			{
				tool: "set_note_flags",
				args: {
					handle: "n3",
					archived: true
				}
			}
		],
		state: [
			{
				note: "n2",
				pinned: true
			},
			{
				note: "n3",
				archived: true
			}
		]
	},
	{
		id: "parallel-favorite-photos-es",
		category: "parallel",
		lang: "es",
		prompt: "Pon IMG_2041.jpg e IMG_2042.jpg en favoritos",
		routes: ["files"],
		calls: [
			{
				tool: "set_favorite",
				args: {
					favorite: true
				}
			}
		],
		state: [
			{
				item: "f11",
				favorite: true
			},
			{
				item: "f12",
				favorite: true
			}
		]
	}
]
