import { describe, expect, it } from "vitest"
import { fileURLToPath } from "node:url"
import { ESLint, type Linter } from "eslint"

const PACKAGE_ROOT = fileURLToPath(new URL("../..", import.meta.url))

// The rule as the real config resolves it for a file, run on its own over a plain-JS fixture: the
// typed config would need the fixture inside the TS project.
async function lintImports(file: string, source: string): Promise<string[]> {
	const config = (await new ESLint({ cwd: PACKAGE_ROOT }).calculateConfigForFile(file)) as Linter.Config
	const rule = config.rules?.["no-restricted-imports"]

	if (rule === undefined) {
		throw new Error(`no import restriction for ${file}`)
	}

	const linter = new ESLint({
		cwd: PACKAGE_ROOT,
		overrideConfigFile: true,
		overrideConfig: [{ files: ["**/*.js"], rules: { "no-restricted-imports": rule } }]
	})
	const [result] = await linter.lintText(source, { filePath: "fixture.js" })

	return (result?.messages ?? []).map(message => message.message)
}

describe("assistant import restrictions", () => {
	it("keeps permanent deletes, unsharing, leaving and logout out of the assistant's reach", async () => {
		const messages = await lintImports(
			"src/features/assistant/runtime/loop.ts",
			[
				'import { deleteItemsPermanently, emptyTrash, trashItems } from "@/features/drive/lib/actions"',
				'import { unshareItems } from "@/features/drive/lib/share/actions"',
				'import { deleteNote, leaveNote } from "@/features/notes/lib/actions"',
				'import { deleteChat } from "@/features/chats/lib/actions"',
				'import { deleteMessage } from "@/features/chats/lib/messageActions"',
				'import { deletePlaylistAction } from "@/features/audio/lib/playlists"',
				'import { blockContact } from "@/features/contacts/lib/actions"',
				'import { requestJobCancel } from "@/features/transfers/lib/control"',
				'import { runLogout } from "@/lib/logout"',
				'import { ChangePassword } from "@/features/settings/components/security/changePassword"'
			].join("\n")
		)

		expect(messages).toHaveLength(12)
		expect(messages.every(message => message.includes("irreversible or security-critical"))).toBe(true)
	})

	it("still allows the trash, which the user can undo", async () => {
		await expect(
			lintImports("src/features/assistant/tools/drive.ts", 'import { trashItems } from "@/features/drive/lib/actions"')
		).resolves.toEqual([])
	})

	it("keeps the tools off the raw SDK worker and relative imports out everywhere", async () => {
		const tools = await lintImports("src/features/assistant/tools/drive.ts", 'import { sdkApi } from "@/lib/sdk/client"\nimport "./x"')

		expect(tools).toHaveLength(2)
		await expect(lintImports("src/features/assistant/runtime/loop.ts", 'import { sdkApi } from "@/lib/sdk/client"')).resolves.toEqual(
			[]
		)
		await expect(
			lintImports("src/features/drive/lib/actions.ts", 'import { emptyTrash } from "@/features/drive/lib/actions"')
		).resolves.toEqual([])
	})
})
