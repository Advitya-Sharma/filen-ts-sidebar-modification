// @vitest-environment happy-dom

import { describe, expect, it, vi } from "vitest"
import { render } from "@testing-library/react"

// The document settings (author name, autosave) follow remote changes in every build with an office editor: one with
// the spreadsheet editor and without DOCX editing too.

const { registerOfficeSettings } = vi.hoisted(() => ({ registerOfficeSettings: vi.fn(() => () => undefined) }))

vi.mock("@/features/preview/lib/docxEditingFlag", () => ({ DOCX_EDITING: false }))
vi.mock("@/features/spreadsheet/next/sheetNextFlag", () => ({ SHEET_NEXT: true }))
vi.mock("@/features/office/lib/officeSettings", () => ({ registerOfficeSettingsSocketHandlers: registerOfficeSettings }))
vi.mock("@/lib/sdk/socket", () => ({ socketBridge: { start: vi.fn(() => Promise.resolve()) } }))
vi.mock("@/features/notes/lib/socketHandlers", () => ({ registerNoteSocketHandlers: () => () => undefined }))
vi.mock("@/features/chats/lib/socketHandlers", () => ({ registerChatSocketHandlers: () => () => undefined }))
vi.mock("@/features/drive/lib/socketHandlers", () => ({ registerDriveSocketHandlers: () => () => undefined }))
vi.mock("@/features/contacts/lib/socketHandlers", () => ({ registerContactSocketHandlers: () => () => undefined }))
vi.mock("@/features/shell/lib/generalSocketHandlers", () => ({ registerGeneralSocketHandlers: () => () => undefined }))
vi.mock("@/features/audio/lib/socketHandlers", () => ({ registerPlaylistSocketHandlers: () => () => undefined }))
vi.mock("@/lib/dotFilen", () => ({ registerDotFilenSocketHandlers: () => () => undefined }))
vi.mock("@/features/videos/lib/library", () => ({ registerLibraryConfigSocketHandlers: () => () => undefined }))
vi.mock("@/features/videos/lib/watchState", () => ({ registerWatchStateSocketHandlers: () => () => undefined }))
vi.mock("@/features/spellcheck/lib/userDictionary", () => ({ registerUserDictionarySocketHandlers: () => () => undefined }))

const { SocketHost } = await import("@/features/shell/components/socketHost")

describe("the socket host", () => {
	it("hears remote changes to the document settings in a build whose only office editor is the spreadsheet's", () => {
		render(<SocketHost />)

		expect(registerOfficeSettings).toHaveBeenCalledTimes(1)
	})
})
