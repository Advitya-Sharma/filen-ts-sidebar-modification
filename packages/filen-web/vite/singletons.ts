// Packages the app must hold exactly one copy of, for resolve.dedupe in vite.config.ts and vitest.config.ts.
export const SINGLETONS = [
	"react",
	"react-dom",
	"@codemirror/state",
	"@codemirror/view",
	"@codemirror/language",
	"@lezer/common",
	"@lezer/highlight",
	"react-markdown",
	"remark-gfm"
]
