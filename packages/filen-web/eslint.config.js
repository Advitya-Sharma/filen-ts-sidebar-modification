import js from "@eslint/js"
import tseslint from "typescript-eslint"
import reactHooks from "eslint-plugin-react-hooks"
import reactRefresh from "eslint-plugin-react-refresh"
import globals from "globals"
import { defineConfig, globalIgnores } from "eslint/config"
import eslintConfigPrettier from "eslint-config-prettier/flat"

const ALIAS_ONLY = {
	group: ["./*", "../*"],
	message: 'Use the "@/..." alias instead of a relative import.'
}

const ASSISTANT_FORBIDDEN =
	"The assistant acts on model output that text other people wrote can steer, so it never reaches irreversible or security-critical operations."

// The assistant's dispatcher runs calls a model produced, and a shared note or chat message can
// steer that model. What the assistant can import it may one day call, so permanent deletion,
// unsharing, leaving, contact removal and blocking, account security, logout and stopping a job
// with trash are kept out of its reach at the import level, whatever the guard decides at runtime.
const ASSISTANT_RESTRICTED_PATHS = [
	{
		name: "@/features/drive/lib/actions",
		importNames: ["deleteItemsPermanently", "emptyTrash", "deleteVersion", "deleteVersions", "disableLink", "disableLinks"],
		message: ASSISTANT_FORBIDDEN
	},
	{ name: "@/features/drive/lib/share/actions", importNames: ["unshareItems"], message: ASSISTANT_FORBIDDEN },
	{ name: "@/features/notes/lib/actions", importNames: ["deleteNote", "leaveNote"], message: ASSISTANT_FORBIDDEN },
	{ name: "@/features/notes/lib/bulk", importNames: ["deleteNotesPermanently", "leaveNotes"], message: ASSISTANT_FORBIDDEN },
	{ name: "@/features/notes/lib/tags", importNames: ["deleteNoteTag"], message: ASSISTANT_FORBIDDEN },
	{ name: "@/features/notes/lib/participants", importNames: ["removeNoteParticipant"], message: ASSISTANT_FORBIDDEN },
	{ name: "@/features/chats/lib/actions", importNames: ["deleteChat", "leaveChat"], message: ASSISTANT_FORBIDDEN },
	{ name: "@/features/chats/lib/bulk", importNames: ["deleteChatsPermanently", "leaveChats"], message: ASSISTANT_FORBIDDEN },
	{ name: "@/features/chats/lib/messageActions", importNames: ["deleteMessage"], message: ASSISTANT_FORBIDDEN },
	{
		name: "@/features/chats/lib/participants",
		importNames: ["removeChatParticipant", "removeChatParticipants"],
		message: ASSISTANT_FORBIDDEN
	},
	{ name: "@/features/audio/lib/playlists", importNames: ["deletePlaylistAction"], message: ASSISTANT_FORBIDDEN },
	{
		name: "@/features/contacts/lib/actions",
		importNames: ["denyRequest", "blockContact", "blockContactByEmail", "toggleParticipantBlocked", "removeContact"],
		message: ASSISTANT_FORBIDDEN
	},
	// requestJobCancel's "trash" mode deletes what a job already wrote.
	{ name: "@/features/transfers/lib/control", importNames: ["requestJobCancel"], message: ASSISTANT_FORBIDDEN },
	{ name: "@/lib/logout", message: ASSISTANT_FORBIDDEN }
]

const ASSISTANT_RESTRICTED_PATTERNS = [
	ALIAS_ONLY,
	{
		group: ["@/features/settings/components/security/*", "@/features/settings/components/account/*", "@/features/auth/**"],
		message: ASSISTANT_FORBIDDEN
	}
]

export default defineConfig([
	globalIgnores([
		"dist",
		"docs",
		"src/routeTree.gen.ts",
		"src/lib/sdk/errorKinds.gen.ts",
		"src/features/settings/thirdPartyNotices.gen.ts"
	]),
	{
		files: ["**/*.{ts,tsx}"],
		extends: [
			js.configs.recommended,
			...tseslint.configs.strictTypeChecked,
			...tseslint.configs.stylisticTypeChecked,
			reactHooks.configs.flat["recommended-latest"], // hooks v7 incl. React Compiler rules
			reactRefresh.configs.vite
		],
		languageOptions: {
			globals: globals.browser,
			parserOptions: { projectService: true, tsconfigRootDir: import.meta.dirname }
		},
		rules: { "@typescript-eslint/no-unused-vars": ["error", { argsIgnorePattern: "^_" }] }
	},
	{
		// shadcn registry output: primitives export a `cva` variants helper
		// alongside the component, which react-refresh can't fast-refresh.
		// Keep these files registry-verbatim instead of splitting them.
		files: ["src/components/ui/**/*.{ts,tsx}"],
		// Also relaxed: some registry files use `Array<T>` / defensive optional chains that our strict
		// config would flag. Scoped to the vendored ui/ dir so app code stays fully strict.
		rules: {
			"react-refresh/only-export-components": "off",
			"@typescript-eslint/array-type": "off",
			"@typescript-eslint/no-unnecessary-condition": "off"
		}
	},
	{
		// TanStack Router file routes export a `Route` constant (not a component) by convention and use
		// `throw redirect()` / `throw notFound()` as the framework's control-flow idiom — neither is a
		// code smell here, so the two rules that would flag them are scoped off for route files.
		files: ["src/routes/**/*.{ts,tsx}"],
		rules: {
			"react-refresh/only-export-components": "off",
			"@typescript-eslint/only-throw-error": "off"
		}
	},
	{
		// The shared unauthed-page guard throws the router's `redirect()` — the same framework
		// control-flow idiom the route files use, extracted so /login and /register share one guard.
		// rootRedirect.ts is the same shape for "/": its beforeLoad body is split out so the redirect
		// switch itself is unit-testable without a router instance (see rootRedirect.test.ts).
		files: ["src/features/auth/lib/guard.ts", "src/features/shell/lib/rootRedirect.ts"],
		rules: {
			"@typescript-eslint/only-throw-error": "off"
		}
	},
	{
		// `@/*` maps to `./src/*` (tsconfig + vite alias + shadcn components.json all agree),
		// so every file under src can reach every other file under src through it — no
		// legitimate relative import can exist here. Scoped to src/ only: vite.config.ts,
		// scripts/*, and vite/* live outside the alias's mapped root and need relative imports.
		files: ["src/**/*.{ts,tsx}"],
		rules: {
			"no-restricted-imports": ["error", { patterns: [ALIAS_ONLY] }]
		}
	},
	{
		files: ["src/features/assistant/**/*.{ts,tsx}"],
		rules: {
			"no-restricted-imports": ["error", { paths: ASSISTANT_RESTRICTED_PATHS, patterns: ASSISTANT_RESTRICTED_PATTERNS }]
		}
	},
	{
		// Tools run every call through the feature modules the user's own actions use, never the raw SDK
		// worker, so the assistant gets the same caches, toasts and confirmations the user does.
		files: ["src/features/assistant/tools/**/*.{ts,tsx}"],
		rules: {
			"no-restricted-imports": [
				"error",
				{
					paths: [
						...ASSISTANT_RESTRICTED_PATHS,
						{ name: "@/lib/sdk/client", message: "Assistant tools go through the feature modules the user's actions use." }
					],
					patterns: ASSISTANT_RESTRICTED_PATTERNS
				}
			]
		}
	},
	{
		// e2e specs are Playwright/Node, not React — the react plugins don't apply, and Playwright's
		// fixture `use` callback trips rules-of-hooks (it reads `use` as the React hook).
		files: ["e2e/**/*.ts"],
		rules: {
			"react-hooks/rules-of-hooks": "off",
			"react-refresh/only-export-components": "off"
		}
	},
	// Must stay last — turns off stylistic rules that would otherwise fight prettier's output.
	eslintConfigPrettier
])
