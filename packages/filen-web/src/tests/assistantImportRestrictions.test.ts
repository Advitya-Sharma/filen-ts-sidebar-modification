import { describe, expect, it } from "vitest"
import { fileURLToPath } from "node:url"
import { ESLint, type Linter } from "eslint"
import tseslint from "typescript-eslint"

const PACKAGE_ROOT = fileURLToPath(new URL("../..", import.meta.url))
const IMPORT_RULES = ["no-restricted-imports", "@typescript-eslint/no-restricted-imports", "no-restricted-syntax"]

// The rules as the real config resolves them for a file, run on their own over an untyped TS
// fixture: the typed config would need the fixture inside the TS project.
async function lintImports(file: string, source: string): Promise<string[]> {
	const config = (await new ESLint({ cwd: PACKAGE_ROOT }).calculateConfigForFile(file)) as Linter.Config
	const rules: Linter.RulesRecord = {}

	for (const name of IMPORT_RULES) {
		const rule = config.rules?.[name]

		if (rule !== undefined) {
			rules[name] = rule
		}
	}

	const linter = new ESLint({
		cwd: PACKAGE_ROOT,
		overrideConfigFile: true,
		overrideConfig: [
			{
				files: ["**/*.ts"],
				languageOptions: { parser: tseslint.parser },
				plugins: { "@typescript-eslint": tseslint.plugin },
				rules
			}
		] as Linter.Config[]
	})
	const [result] = await linter.lintText(source, { filePath: "fixture.ts" })

	return (result?.messages ?? []).map(message => message.message)
}

const FORBIDDEN = [
	'import { deleteItemsPermanently, emptyTrash, trashItems } from "@/features/drive/lib/actions"',
	'import { createLink, updateLink } from "@/features/drive/lib/actions"',
	'import { shareItems, unshareItems } from "@/features/drive/lib/share/actions"',
	'import { deleteNote, leaveNote } from "@/features/notes/lib/actions"',
	'import { addNoteParticipants, setNoteParticipantPermission } from "@/features/notes/lib/participants"',
	'import { deleteChat } from "@/features/chats/lib/actions"',
	'import { enqueueChatMessage, sync } from "@/features/chats/lib/sync"',
	'import { deleteMessage, editMessage, disableMessageEmbed } from "@/features/chats/lib/messageActions"',
	'import { deleteDraft } from "@/features/chats/lib/drafts"',
	'import { addChatParticipants } from "@/features/chats/lib/participants"',
	'import { retryInflightMessage, purgeChatInflightState } from "@/features/chats/lib/inflight"',
	'import { startChatOutbox } from "@/features/chats/lib/outboxCoordinator"',
	'import { trustDomain } from "@/features/chats/lib/trustedDomains"',
	'import { uploadAttachment } from "@/features/chats/lib/attachments"',
	'import { enqueueInflightMessage, useChatsInflightStore } from "@/features/chats/store/useChatsInflight"',
	'import { deletePlaylistAction } from "@/features/audio/lib/playlists"',
	'import { blockContact, sendContactRequest, acceptRequest, messageContact } from "@/features/contacts/lib/actions"',
	'import { requestJobCancel } from "@/features/transfers/lib/control"',
	'import { cancelTransfers, setTransfersPaused } from "@/features/transfers/lib/control"',
	'import { restoreItems, restoreVersion } from "@/features/drive/lib/actions"',
	'import { restoreNote } from "@/features/notes/lib/actions"',
	'import { restoreNotes } from "@/features/notes/lib/bulk"',
	'import { restoreNoteFromHistory } from "@/features/notes/lib/history"',
	'import { clearSession, persistSession, resumeSession, broadcastAuth } from "@/lib/sdk/session"',
	'import { runLogout } from "@/lib/logout"',
	'import { performLogout } from "@/features/shell/lib/performLogout"',
	'import { ChangePassword } from "@/features/settings/components/security/changePassword"',
	'import { setTheme } from "@/features/settings/lib/theme"',
	'import { pendingSdkCalls } from "@/e2e-hooks/sdkCalls"'
]

describe("assistant import restrictions", () => {
	it.each(["src/features/assistant/runtime/loop.ts", "src/features/assistant/tools/drive.ts", "src/workers/assistant.worker.ts"])(
		"keeps deleting, restoring, sending, queueing, sharing, settings, transfer control, the session and logout out of reach in %s",
		async file => {
			const messages = await lintImports(file, FORBIDDEN.join("\n"))

			// One per banned name, or per statement for a module banned whole; trashItems passes.
			expect(messages).toHaveLength(36)
			expect(messages.every(message => message.includes("irreversible or security-critical"))).toBe(true)

			for (const line of FORBIDDEN) {
				await expect(lintImports(file, line)).resolves.not.toEqual([])
			}
		}
	)

	it("still allows the trash, which the user can undo, and reading a draft", async () => {
		await expect(
			lintImports(
				"src/features/assistant/tools/drive.ts",
				'import { trashItems } from "@/features/drive/lib/actions"\nimport { loadDraft } from "@/features/chats/lib/drafts"'
			)
		).resolves.toEqual([])
	})

	it.each(["src/features/assistant/runtime/loop.ts", "src/features/assistant/store/x.ts", "src/features/assistant/tools/drive.ts"])(
		"keeps %s off the raw SDK, allowing only its types",
		async file => {
			const messages = await lintImports(
				file,
				[
					'import { sdkApi } from "@/lib/sdk/client"',
					'import { api } from "@/workers/sdk.worker"',
					'import { FilenSdkError } from "@filen/sdk-rs"'
				].join("\n")
			)

			expect(messages).toHaveLength(3)
			expect(messages.every(message => message.includes("never the raw SDK"))).toBe(true)
			await expect(
				lintImports(
					file,
					'import type { Note } from "@filen/sdk-rs"\nimport { type Chat } from "@filen/sdk-rs"\nimport type { ArchiveNameInfo } from "@/workers/sdk.worker"'
				)
			).resolves.toEqual([])
		}
	)

	it("refuses a dynamic import, which the import rules cannot see", async () => {
		await expect(lintImports("src/features/assistant/loop/x.ts", 'void import("@/lib/sdk/client")')).resolves.toEqual([
			"A dynamic import would slip past the assistant's import restrictions."
		])
	})

	it("keeps relative imports out everywhere and leaves the rest of the app alone", async () => {
		await expect(lintImports("src/features/assistant/tools/drive.ts", 'import "./x"')).resolves.toHaveLength(1)
		await expect(
			lintImports(
				"src/features/drive/lib/actions.ts",
				'import { emptyTrash } from "@/features/drive/lib/actions"\nimport { sdkApi } from "@/lib/sdk/client"\nvoid import("@/lib/sdk/client")'
			)
		).resolves.toEqual(["No dynamic imports in app code; import statically."])
	})
})
