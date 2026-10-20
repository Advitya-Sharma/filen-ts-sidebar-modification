// Packages the app must hold exactly one copy of, for resolve.dedupe in vite.config.ts and vitest.config.ts.
export const SINGLETONS = [
	"react",
	"react-dom",
	"@codemirror/state",
	"@codemirror/view",
	"@codemirror/language",
	"@codemirror/autocomplete",
	"@codemirror/commands",
	"@codemirror/lint",
	"@codemirror/merge",
	"@codemirror/search",
	"@lezer/common",
	"@lezer/highlight",
	"react-markdown",
	"remark-gfm",
	"pdfjs-dist"
]
