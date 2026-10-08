// English source catalog — "assistant" namespace: the on-device assistant panel — its setup and download,
// the conversation, tool activity, confirmation cards and the composer. Same typed-catalog rules as every
// other namespace: flat `as const` object, camelCase keys, no literal '.' or ':' (real i18next
// namespaces, keySeparator/nsSeparator both ON). "directory" never "folder".
export const assistant = {
	// ── Panel chrome ─────────────────────────────────────────────────────────
	/** Panel heading and the accessible name of the panel landmark and its narrow-screen drawer */
	assistantTitle: "Assistant",
	/** Header button — closes the panel */
	assistantClose: "Close assistant",
	/** Header button — clears the conversation and starts over */
	assistantNewConversation: "New conversation",
	/** Header button — opens the overflow menu */
	assistantMoreOptions: "More options",
	/** Overflow menu — deletes the downloaded model from this device */
	assistantRemoveModel: "Remove downloaded model",
	/** Remove-model confirmation — title */
	assistantRemoveModelTitle: "Remove the downloaded model?",
	/** Remove-model confirmation — body; frees disk space, can be downloaded again */
	assistantRemoveModelBody: "This frees {{size}} on this device. You can download the model again at any time.",
	/** Remove-model confirmation — confirm button */
	assistantRemoveModelConfirm: "Remove",
	/** Toast — removing the model failed */
	assistantRemoveModelFailed: "Couldn't remove the downloaded model",
	/** Drag handle on the panel's left edge — accessible name of the separator */
	assistantResize: "Resize assistant panel",

	// ── Status chip in the header ────────────────────────────────────────────
	/** Status chip — the device can't run the model */
	assistantStatusUnavailable: "Unavailable",
	/** Status chip — another Filen tab has the model loaded */
	assistantStatusInUseElsewhere: "In another tab",
	/** Status chip — the model hasn't been downloaded yet */
	assistantStatusNotDownloaded: "Not set up",
	/** Status chip — download progress, a whole percentage */
	assistantStatusDownloading: "Downloading {{percent}}%",
	/** Status chip — the downloaded model is being started */
	assistantStatusLoading: "Starting",
	/** Status chip — ready, and running on this device */
	assistantStatusReady: "On this device",
	/** Status chip — downloading or starting the model failed */
	assistantStatusError: "Error",

	// ── Setup (first run) ────────────────────────────────────────────────────
	/** Setup screen — title */
	assistantSetupTitle: "A private assistant on this device",
	/** Setup screen — what it is and the privacy promise: the model runs locally, nothing is sent anywhere */
	assistantSetupBody:
		"Ask it to find, open and summarize your files, notes and chats. It runs entirely on this device: your data and your questions are never sent to Filen or anyone else.",
	/** Setup screen — label of the recommended model's name */
	assistantSetupModel: "Model",
	/** Setup screen — label of the one-time download size */
	assistantSetupSize: "Download",
	/** Setup screen — the model is stored locally and survives signing out */
	assistantSetupStorageNote: "Downloaded once and kept on this device, also after you sign out.",
	/** Setup screen — starts the model download */
	assistantDownload: "Download model",
	/** Setup screen — the device check failed; the download button stays disabled below this */
	assistantUnavailableTitle: "The assistant can't run here",
	/** Unavailable reason — the browser lacks what the model needs */
	assistantUnavailableUnsupportedBrowser:
		"This browser can't run the assistant's model. Try a recent version of Chrome, Edge or Safari on a computer.",
	/** Unavailable reason — not enough memory */
	assistantUnavailableInsufficientMemory: "This device doesn't have enough memory to run the assistant's model.",
	/** Unavailable reason — not enough free storage for the download */
	assistantUnavailableInsufficientStorage: "There isn't enough free storage on this device for the assistant's model.",
	/** Another Filen tab holds the model; it becomes available here once that tab closes it — title */
	assistantInUseElsewhereTitle: "Open in another Filen tab",
	/** Another Filen tab holds the model — explanation; no action is offered */
	assistantInUseElsewhereBody:
		"The assistant runs in one tab at a time. Close it in the other tab, or use it there; it becomes available here on its own.",

	// ── Download and start ───────────────────────────────────────────────────
	/** Download in progress — title */
	assistantDownloadingTitle: "Downloading the model",
	/** Download in progress — bytes received of the total, e.g. "1.2 GiB of 2.6 GiB" */
	assistantDownloadProgress: "{{received}} of {{total}}",
	/** Download in progress — offline, or retrying after the connection dropped; it continues on its own */
	assistantDownloadWaiting: "Waiting for the connection…",
	/** Download in progress — every byte has arrived, the last checks are running */
	assistantDownloadFinishing: "Finishing the download…",
	/** Download in progress — accessible name of the progress bar */
	assistantDownloadProgressLabel: "Model download",
	/** Download in progress — stops it */
	assistantCancelDownload: "Cancel download",
	/** The downloaded model is being loaded into memory */
	assistantLoadingTitle: "Starting the assistant…",
	/** Before starting, the downloaded model's files are checked because a record of them was lost or the last start failed */
	assistantVerifyingTitle: "Checking the downloaded model…",
	/** Download or start failed — title above the reason */
	assistantErrorTitle: "The assistant couldn't start",
	/** Failure reason — the download broke off */
	assistantErrorDownloadFailed: "The download failed. Check your connection and try again.",
	/** Failure reason — the model couldn't be loaded */
	assistantErrorLoadFailed: "The model couldn't be started.",
	/** Failure reason — out of memory while loading */
	assistantErrorOutOfMemory: "This device ran out of memory while starting the model. Close other tabs and try again.",
	/** Failure reason — storage filled up during the download */
	assistantErrorStorageFull: "This device ran out of storage while downloading the model.",
	/** Failure — retries the download or start */
	assistantRetry: "Try again",

	// ── Conversation ─────────────────────────────────────────────────────────
	/** Empty conversation — title */
	assistantEmptyTitle: "Ask about your files, notes and chats",
	/** Empty conversation — subtitle */
	assistantEmptyBody: "Everything stays on this device. Try one of these:",
	/** Empty conversation — accessible name of the example prompt group */
	assistantExamplesLabel: "Example questions",
	/** Example prompt — sent as typed when clicked */
	assistantExampleFind: "Find the PDFs I uploaded this month",
	/** Example prompt — sent as typed when clicked */
	assistantExampleSummarize: "Summarize my most recent note",
	/** Example prompt — sent as typed when clicked */
	assistantExampleOpen: "Open my Documents directory",
	/** Example prompt — sent as typed when clicked */
	assistantExamplePlay: "Play my most recent playlist",
	/** Conversation scroller — accessible name */
	assistantConversationLabel: "Conversation",
	/** Screen-reader prefix on the user's own messages */
	assistantYouSaid: "You",
	/** Accessible name of one assistant reply */
	assistantReplyLabel: "Assistant reply",
	/** Announced once when a reply starts and shown as the waiting bubble's label */
	assistantResponding: "Assistant is responding…",
	/** Screen-reader suffix on a tool activity that failed */
	assistantActivityFailed: "Failed",
	/** Turn ended — the user pressed Stop */
	assistantTurnCancelled: "Stopped",
	/** Turn failed — the model isn't loaded */
	assistantTurnErrorModelUnavailable: "The model isn't ready.",
	/** Turn failed — the conversation outgrew the model's context */
	assistantTurnErrorContextFull: "This conversation is too long. Start a new one to continue.",
	/** Turn failed — anything else */
	assistantTurnErrorFailed: "Something went wrong while answering.",

	// ── Tool activity (one line per tool call, shown while it runs and after) ─
	/** Activity — get_context */
	assistantToolGetContext: "Looking at your screen",
	/** Activity — search_items */
	assistantToolSearchItems: "Searching your drive",
	/** Activity — list_directory */
	assistantToolListDirectory: "Listing a directory",
	/** Activity — get_item_info */
	assistantToolGetItemInfo: "Checking item details",
	/** Activity — read_text_file */
	assistantToolReadTextFile: "Reading a file",
	/** Activity — reveal_item */
	assistantToolRevealItem: "Showing an item in your drive",
	/** Activity — set_favorite */
	assistantToolSetFavorite: "Updating favorites",
	/** Activity — create_directory */
	assistantToolCreateDirectory: "Creating a directory",
	/** Activity — rename_item */
	assistantToolRenameItem: "Renaming an item",
	/** Activity — move_items */
	assistantToolMoveItems: "Moving items",
	/** Activity — copy_items */
	assistantToolCopyItems: "Copying items",
	/** Activity — trash_items */
	assistantToolTrashItems: "Moving items to the trash",
	/** Activity — list_notes */
	assistantToolListNotes: "Listing your notes",
	/** Activity — read_note */
	assistantToolReadNote: "Reading a note",
	/** Activity — open_note */
	assistantToolOpenNote: "Opening a note",
	/** Activity — create_note */
	assistantToolCreateNote: "Creating a note",
	/** Activity — set_note_flags */
	assistantToolSetNoteFlags: "Updating a note",
	/** Activity — list_chats */
	assistantToolListChats: "Listing your chats",
	/** Activity — read_chat */
	assistantToolReadChat: "Reading a chat",
	/** Activity — open_chat */
	assistantToolOpenChat: "Opening a chat",
	/** Activity — draft_chat_message */
	assistantToolDraftChatMessage: "Drafting a message",
	/** Activity — list_playlists */
	assistantToolListPlaylists: "Listing your playlists",
	/** Activity — open_playlist */
	assistantToolOpenPlaylist: "Opening a playlist",
	/** Activity — play */
	assistantToolPlay: "Starting playback",
	/** Activity — audio_control */
	assistantToolAudioControl: "Controlling playback",
	/** Activity — add_tracks_to_playlist */
	assistantToolAddTracksToPlaylist: "Adding tracks to a playlist",
	/** Activity — open_route */
	assistantToolOpenRoute: "Opening a screen",
	/** Activity — a tool this version has no wording for */
	assistantToolGeneric: "Working",

	// ── Confirmation cards ───────────────────────────────────────────────────
	/** Card title — a drafted chat message, before it goes into the chat's input field */
	assistantProposalDraftMessage: "Put this draft into {{chat}}?",
	/** Draft card — the user still sends the message; nothing is sent automatically */
	assistantProposalDraftNote: "You still send it yourself.",
	/** Draft card — accessible name of the drafted message text */
	assistantProposalDraftText: "Draft",
	/** Draft card — confirm button */
	assistantConfirmDraft: "Add draft",
	/** Card title — move */
	assistantProposalMove_one: "Move {{count}} item",
	/** Card title — move */
	assistantProposalMove_other: "Move {{count}} items",
	/** Card title — copy */
	assistantProposalCopy_one: "Copy {{count}} item",
	/** Card title — copy */
	assistantProposalCopy_other: "Copy {{count}} items",
	/** Card title — trash */
	assistantProposalTrash_one: "Move {{count}} item to the trash",
	/** Card title — trash */
	assistantProposalTrash_other: "Move {{count}} items to the trash",
	/** Card title — download */
	assistantProposalDownload_one: "Download {{count}} item",
	/** Card title — download */
	assistantProposalDownload_other: "Download {{count}} items",
	/** Card title — rename */
	assistantProposalRename_one: "Rename {{count}} item",
	/** Card title — rename */
	assistantProposalRename_other: "Rename {{count}} items",
	/** Card title — add to favorites */
	assistantProposalFavorite_one: "Add {{count}} item to favorites",
	/** Card title — add to favorites */
	assistantProposalFavorite_other: "Add {{count}} items to favorites",
	/** Card title — remove from favorites */
	assistantProposalUnfavorite_one: "Remove {{count}} item from favorites",
	/** Card title — remove from favorites */
	assistantProposalUnfavorite_other: "Remove {{count}} items from favorites",
	/** Card title — create a directory */
	assistantProposalCreateDirectory: "Create a directory",
	/** Card title — create a note */
	assistantProposalCreateNote: "Create a note",
	/** Card field — the target directory of a move, copy or create */
	assistantProposalDestination: "In",
	/** Card field — the new name of a rename or create */
	assistantProposalName: "Name",
	/** Card field — any other free-text argument */
	assistantProposalText: "Text",
	/** Card — accessible name of the affected-items list */
	assistantProposalItems: "Items",
	/** Card — runs the action */
	assistantConfirm: "Confirm",
	/** Card — runs a trash action */
	assistantConfirmTrash: "Move to trash",
	/** Card — declines the action */
	assistantDecline: "Cancel",
	/** Card status — confirmed, in progress */
	assistantProposalRunning: "Working…",
	/** Card status — finished */
	assistantProposalDone: "Done",
	/** Card status — failed */
	assistantProposalFailed: "Couldn't finish this",
	/** Card status — declined */
	assistantProposalCancelled: "Cancelled",
	/** Card status — a newer message was sent before it was answered */
	assistantProposalExpired: "Expired. Ask again to do this.",
	/** Announced when a card appears, after the reply */
	assistantProposalAnnouncement: "Confirmation needed: {{title}}",

	// ── Composer ─────────────────────────────────────────────────────────────
	/** Composer — placeholder and accessible name */
	assistantComposerPlaceholder: "Ask the assistant",
	/** Composer — send button */
	assistantSend: "Send",
	/** Composer — stops the reply being generated */
	assistantStop: "Stop"
} as const
