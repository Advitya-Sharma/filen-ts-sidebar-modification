// English source catalog — "videos" namespace: the /videos screens, the Videos sidebar and phone tabs, the
// directory chooser and the Libraries dialog. Same typed-catalog rules as every namespace: flat `as const`
// object, camelCase keys, no literal '.' or ':'. The module's own name is common:moduleVideos.
export const videos = {
	// ── Navigation ─────────────────────────────────────────────────────────────
	/** Sidebar + phone tabs — the library's home */
	videosHome: "Home",
	/** Sidebar + phone tabs — every movie in the library */
	videosMovies: "Movies",
	/** Sidebar + phone tabs — every series in the library */
	videosShows: "Shows",
	/** Sidebar — videos that are neither a movie nor an episode; listed only when there are some */
	videosOther: "Other",
	/** Phone tabs — searching the library */
	videosSearch: "Search",
	/** Sidebar — header over the list of library directories */
	videosLibrariesGroup: "Libraries",
	/** Sidebar footer, header button — opens the Libraries dialog */
	videosManageLibraries: "Manage libraries",
	/** Sidebar header icon button, first-run button, Libraries dialog footer — opens the directory chooser */
	videosAddDirectory: "Add a directory",
	/** Sidebar — accessible label of the drag handle that resizes the panel */
	videosSidebarResize: "Resize sidebar",
	/** Sidebar library row — screen-reader text replacing the count while the directory is scanned */
	videosRootScanning: "Scanning",
	/** Sidebar library row — screen-reader text replacing the count when the latest scan didn't finish */
	videosRootIncomplete: "Incomplete",
	/** Sidebar library row — screen-reader text replacing the count when the directory can't be found */
	videosRootMissing: "Not found",
	/** Sidebar row — screen-reader form of a count shown as a bare number (singular) */
	videosCountTitles_one: "{{count, number}} title",
	/** Sidebar row — screen-reader form of a count shown as a bare number (plural) */
	videosCountTitles_other: "{{count, number}} titles",

	// ── First run ──────────────────────────────────────────────────────────────
	/** First run — title before any directory is part of the library */
	videosFirstRunTitle: "Add a directory to build your library",
	/** First run — what the library is made from */
	videosFirstRunBody: "Titles come from file and directory names. Nothing is uploaded or sent anywhere.",
	/** First run — hint under the button */
	videosFirstRunHint: "You can add more directories later, for example one for movies and one for shows.",
	/** Every directory is turned off in Libraries — title */
	videosAllOffTitle: "No directory is included",
	/** Every directory is turned off in Libraries — body */
	videosAllOffBody: "Turn one on in Libraries, or add another directory.",

	// ── Scanning ───────────────────────────────────────────────────────────────
	/** Scan status line — one directory being scanned; {{name}} = directory name. No count: a scan is one request whose result arrives at once */
	videosScanning: "Scanning “{{name}}”…",
	/** Scan status line — several directories being scanned (singular remainder) */
	videosScanningMore_one: "Scanning “{{name}}” and {{count, number}} more…",
	/** Scan status line — several directories being scanned (plural remainder) */
	videosScanningMore_other: "Scanning “{{name}}” and {{count, number}} more…",
	/** Scan status line — button that stops showing the scan; the previous library stays */
	videosStopScan: "Stop",
	/** Scan status line — accessible label of the Stop button */
	videosStopScanLabel: "Stop scanning",
	/** Scan status line — retries directories whose scan never started */
	videosRetryScan: "Retry",
	/** Screen-reader announcement once a directory's scan finished (singular) */
	videosScanDone_one: "{{name}} scanned: {{count, number}} video",
	/** Screen-reader announcement once a directory's scan finished (plural) */
	videosScanDone_other: "{{name}} scanned: {{count, number}} videos",
	/** Empty library — the only directory holds no video files; {{name}} = directory name */
	videosRootEmptyTitle: "No video files found in {{name}}",
	/** Empty library — several directories, none holding video files */
	videosLibraryEmptyTitle: "No video files found",
	/** Empty library — body */
	videosLibraryEmptyBody: "Video files anywhere inside a library directory, including its subdirectories, show up here.",
	/** Button — scans a directory again */
	videosRescan: "Rescan",
	/** Empty library — opens the chooser to replace the empty directory with another */
	videosChooseAnother: "Choose another",
	/** Banner — a directory's latest scan didn't finish; an earlier complete scan is shown; {{name}} = directory name */
	videosIncompleteBanner: "{{name}}: library incomplete. Showing the last complete scan.",
	/** Banner — a directory's first scan didn't finish, so nothing of it is shown */
	videosIncompleteBannerNoScan: "{{name}}: library incomplete. The scan didn't finish.",
	/** Banner — a library directory can't be found any more */
	videosMissingBanner: "{{name}} can't be found. It may have been moved to the trash or deleted.",

	// ── Libraries dialog ───────────────────────────────────────────────────────
	/** Libraries dialog — title */
	videosLibrariesTitle: "Libraries",
	/** Libraries dialog — description under the title */
	videosLibrariesDescription: "Each directory is read by name only. Watch progress syncs to your account.",
	/** Libraries dialog — shown when the library has no directory yet */
	videosLibrariesNone: "No directory is part of your library yet.",
	/** Libraries dialog — accessible label of a directory's include switch */
	videosLibraryInclude: "Include {{name}} in Videos",
	/** Libraries dialog — label of a directory's type select */
	videosLibraryType: "Type",
	/** Type select — classify each title by its names */
	videosTypeAuto: "Auto",
	/** Type select — auto, with what the names were classified as; {{kind}} = Movies, Shows or Other */
	videosTypeAutoDetected: "Auto: {{kind}}",
	/** Type select — every title is a movie */
	videosTypeMovies: "Movies",
	/** Type select — every title is a series */
	videosTypeShows: "Shows",
	/** Type select — no movie or series matching */
	videosTypeOther: "Other",
	/** Libraries dialog subline — video count (singular) */
	videosLibraryVideos_one: "{{count, number}} video",
	/** Libraries dialog subline — video count (plural) */
	videosLibraryVideos_other: "{{count, number}} videos",
	/** Libraries dialog subline — series count (singular) */
	videosLibraryShows_one: "{{count, number}} show",
	/** Libraries dialog subline — series count (plural) */
	videosLibraryShows_other: "{{count, number}} shows",
	/** Libraries dialog subline — path, counts and scan; {{path}} = directory path, {{counts}} and {{scan}} = the parts below */
	videosLibrarySubline: "{{path}} · {{counts}} · {{scan}}",
	/** Libraries dialog subline — a turned-off directory */
	videosLibrarySublineOff: "{{path}} · not included",
	/** Libraries dialog subline — video and series counts together */
	videosLibraryCountsWithShows: "{{videos}} · {{shows}}",
	/** Libraries dialog subline — when the last complete scan began; {{time}} = relative time */
	videosLibraryScanned: "last scan: {{time}}",
	/** Libraries dialog subline — the directory was never scanned completely */
	videosLibraryNotScanned: "not scanned yet",
	/** Libraries dialog — the directory is being scanned */
	videosLibraryScanning: "Scanning",
	/** Libraries dialog — the latest scan didn't finish, an earlier one is shown */
	videosLibraryIncomplete: "Library incomplete: the last scan didn't finish. Showing the previous one.",
	/** Libraries dialog — the first scan didn't finish */
	videosLibraryIncompleteNoScan: "Library incomplete: the scan didn't finish.",
	/** Libraries dialog — the directory can't be found */
	videosLibraryMissing: "Not found. It may have been moved to the trash or deleted.",
	/** Libraries dialog — the directory was moved since it was added */
	videosLibraryMoved: "Moved since it was added.",
	/** Libraries dialog — a large directory; {{files, number}} = files and directories it holds */
	videosLibrarySizeLarge: "About {{files, number}} files. Scanning it uses a lot of memory.",
	/** Libraries dialog — a very large directory */
	videosLibrarySizeHuge: "About {{files, number}} files. Scanning it can run out of memory; a narrower directory is lighter.",
	/** Libraries dialog — accessible label of a directory's menu trigger */
	videosLibraryMenu: "More options for {{name}}",
	/** Libraries dialog menu — opens the directory in Cloud Drive */
	videosOpenInDrive: "Open in Cloud Drive",
	/** Libraries dialog menu — removes the directory from Videos */
	videosRemoveRoot: "Remove from Videos",
	/** Remove confirmation — title */
	videosRemoveConfirmTitle: "Remove {{name}}?",
	/** Remove confirmation — body */
	videosRemoveConfirmBody: "Removes {{name}} from Videos only. Files and watch progress stay.",
	/** Remove confirmation — confirm button */
	videosRemoveConfirmAction: "Remove",
	/** Libraries dialog — footer button closing it */
	videosLibrariesDone: "Done",

	// ── Directory chooser ──────────────────────────────────────────────────────
	/** Chooser — title */
	videosChooserTitle: "Add a directory to Videos",
	/** Chooser — filter input placeholder and accessible label */
	videosChooserFilterPlaceholder: "Filter directories",
	/** Chooser — confirm button adding the open directory; {{name}} = its name */
	videosChooserConfirm: "Add “{{name}}”",
	/** Chooser row — the directory is already a library directory */
	videosChooserAlreadyRoot: "Already in Videos",
	/** Chooser — the open directory sits inside a library directory; {{name}} = that directory */
	videosChooserInsideRoot: "Inside {{name}}",
	/** Chooser — the open directory overlaps library directories; {{names}} = their names */
	videosChooserReplaceBody: "Library directories can't overlap. Adding {{name}} replaces {{names}}.",
	/** Chooser — confirm a replacement */
	videosChooserReplaceAction: "Replace",
	/** Chooser — the open directory is Filen's own */
	videosChooserSystemDirectory: "This directory belongs to Filen and can't be added.",
	/** Chooser — the open directory can't be found any more */
	videosChooserNotFound: "This directory can't be found any more.",
	/** Chooser — the open directory is very large; {{files, number}} = files and directories inside */
	videosChooserLargeBody:
		"{{name}} holds about {{files, number}} files. Scanning it uses a lot of memory; a narrower directory, such as one for movies and one for shows, is lighter.",
	/** Chooser — add a large directory anyway */
	videosChooserLargeAction: "Add anyway",
	/** Chooser — leaves a warning and goes back to browsing */
	videosChooserBack: "Back"
} as const
