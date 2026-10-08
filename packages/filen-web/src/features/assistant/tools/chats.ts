import type { Chat } from "@filen/sdk-rs"
import { sanitizeUntrustedText } from "@filen/shared/assistant"
import { chatsQueryRead } from "@/features/chats/queries/chats"
import { chatMessagesQueryRead } from "@/features/chats/queries/chatMessages"
import { chatsWithoutBlockedOneOnOne, filterChats } from "@/features/chats/components/chatsSidebar.logic"
import { isChatUndecryptable, messageSenderName, otherParticipants } from "@/features/chats/lib/sort"
import { isOwnMessage, isSenderBlocked } from "@/features/chats/lib/sender"
import { isOverLimit } from "@/features/chats/lib/composer.logic"
import { loadDraft } from "@/features/chats/lib/drafts"
import { useChatComposerStore } from "@/features/chats/store/useChatComposer"
import { i18n } from "@/lib/i18n"
import { assistantRouter, navigateChecked } from "@/features/assistant/tools/navigation"
import {
	blockedUsers,
	chatName,
	currentUserId,
	entityVisible,
	freshChat,
	optionalString,
	participantName,
	requiredString,
	resolveHandle,
	type ToolAdapter,
	type ToolScope
} from "@/features/assistant/tools/entities"
import { AssistantToolError, clipText, isoDate, matchScore, searchTokens } from "@/features/assistant/tools/format"

const LIST_LIMIT = 30
const DEFAULT_MESSAGES = 20
const MAX_MESSAGES = 50
const MESSAGE_MAX_CHARS = 600
export const CHAT_MAX_CHARS = 6_000

async function resolveChat(scope: ToolScope, handle: string): Promise<Chat> {
	const entity = await resolveHandle(scope, handle, ["chat"])

	if (entity.kind !== "chat") {
		throw new AssistantToolError("Not a chat", { invalid: true })
	}

	const chat = freshChat(entity.chat)

	if (!entityVisible({ kind: "chat", chat })) {
		throw new AssistantToolError("This chat is not available")
	}

	if (isChatUndecryptable(chat)) {
		throw new AssistantToolError("This chat cannot be decrypted")
	}

	return chat
}

// The sidebar's own list, in its order; a query matches the chat's name and its participants' names word by
// word, both sides folded the same way ("joao" finds "João"), the chats matching every word first.
export const listChatsTool: ToolAdapter = async (input, scope) => {
	const userId = currentUserId()
	const blocked = blockedUsers()
	const visible = chatsWithoutBlockedOneOnOne(await chatsQueryRead(), userId, blocked)
	const tokens = searchTokens(optionalString(input, "query") ?? "")
	const listed = filterChats(visible, "", userId, i18n.t("chats:chatJustYou"), blocked)
	const chats =
		tokens.length === 0
			? listed
			: listed
					.map(chat => ({
						chat,
						score: matchScore([chatName(chat), ...otherParticipants(chat, userId).map(participantName)].join(" "), tokens)
					}))
					.filter(({ score }) => score > 0)
					.toSorted((a, b) => b.score - a.score)
					.map(({ chat }) => chat)

	return {
		value: {
			chats: chats.slice(0, LIST_LIMIT).map(chat => {
				const name = chatName(chat)

				return {
					handle: scope.show({ kind: "chat", chat }, name),
					name,
					participants: otherParticipants(chat, userId).map(participantName),
					...(chat.lastMessage === undefined ? {} : { lastActivity: isoDate(chat.lastMessage.sentTimestamp) })
				}
			}),
			...(chats.length > LIST_LIMIT ? { more: chats.length - LIST_LIMIT } : {})
		}
	}
}

// The thread's newest messages from its own cache, as the open thread shows them: a blocked sender's
// text stays hidden. Reading never marks the chat read.
export const readChatTool: ToolAdapter = async (input, scope) => {
	const chat = await resolveChat(scope, requiredString(input, "handle"))
	const requested = typeof input["limit"] === "number" && input["limit"] > 0 ? Math.floor(input["limit"]) : DEFAULT_MESSAGES
	const limit = Math.min(requested, MAX_MESSAGES)
	const userId = currentUserId()
	const blocked = blockedUsers()
	const name = chatName(chat)
	const messages = (await chatMessagesQueryRead(chat.uuid)).slice(-limit)
	let budget = CHAT_MAX_CHARS
	let truncated = false
	const rows: { from: string; text: string; sent: string }[] = []

	// Newest first against the budget, so a long backlog drops its oldest messages.
	for (const message of [...messages].reverse()) {
		const text = isSenderBlocked(message, blocked)
			? "(hidden: blocked sender)"
			: message.message === undefined
				? "(cannot decrypt)"
				: clipText(message.message, Math.min(MESSAGE_MAX_CHARS, budget)).text

		if (budget <= 0) {
			truncated = true
			break
		}

		budget -= text.length
		rows.unshift({
			from: isOwnMessage(message, userId) ? "you" : messageSenderName(message),
			text,
			sent: isoDate(message.sentTimestamp)
		})
	}

	return {
		value: {
			chat: { handle: scope.show({ kind: "chat", chat }, name), name },
			messages: rows,
			...(truncated ? { truncated: true } : {})
		}
	}
}

export const openChatTool: ToolAdapter = async (input, scope) => {
	const chat = await resolveChat(scope, requiredString(input, "handle"))
	const router = assistantRouter()

	if (router === null) {
		throw new AssistantToolError("The app cannot navigate right now")
	}

	await navigateChecked({
		router,
		navigate: () => router.navigate({ to: "/chats/$uuid", params: { uuid: chat.uuid } }),
		landed: pathname => pathname === `/chats/${chat.uuid}`,
		signal: scope.signal
	})

	return { value: { opened: chatName(chat) }, uiEffect: { type: "navigated", pathname: `/chats/${chat.uuid}` } }
}

// The text a draft puts into the composer: what the card shows and what lands, with invisible and bidi
// characters and template control tokens removed, so the card cannot show one message and send another.
export function draftText(text: string): string {
	return sanitizeUntrustedText(text)
}

// Why a draft cannot land, checked before the card shows (and again after the click, for a race): the
// user is never asked to confirm a draft that would then fail.
export async function draftProblem({ chat, text }: { chat: Chat; text: string }): Promise<string | null> {
	if (text.trim().length === 0) {
		return "The message is empty"
	}

	if (isOverLimit(text)) {
		return "The message is too long"
	}

	const entry = useChatComposerStore.getState().entries[chat.uuid]
	const saved = entry === undefined ? await loadDraft(chat.uuid) : ""

	if ((entry !== undefined && (entry.draft.trim().length > 0 || entry.mode.kind !== "new")) || saved.trim().length > 0) {
		return `The message box of "${chatName(chat)}" already holds unsent text; the user has to clear it first`
	}

	return null
}

export async function draftChatProblem(scope: ToolScope, input: Record<string, unknown>): Promise<string | null> {
	return draftProblem({ chat: await resolveChat(scope, requiredString(input, "chat")), text: draftText(requiredString(input, "text")) })
}

// Runs only after the user confirmed the draft. The text goes into the composer exactly as the user's
// own typing would (its store, mirrored to disk by the composer) and the thread opens; the user sends it.
// Text the user already left in that composer is never overwritten.
export const draftChatMessageTool: ToolAdapter = async (input, scope) => {
	const chat = await resolveChat(scope, requiredString(input, "chat"))
	const text = draftText(requiredString(input, "text"))
	const name = chatName(chat)
	const router = assistantRouter()

	if (router === null) {
		throw new AssistantToolError("The app cannot navigate right now")
	}

	const problem = await draftProblem({ chat, text })

	if (problem !== null) {
		throw new AssistantToolError(problem)
	}

	// The draft is written only once the thread is open: a navigation a dialog refused leaves no draft behind.
	await navigateChecked({
		router,
		navigate: () => router.navigate({ to: "/chats/$uuid", params: { uuid: chat.uuid } }),
		landed: pathname => pathname === `/chats/${chat.uuid}`,
		signal: scope.signal
	})

	const raced = await draftProblem({ chat, text })

	if (raced !== null) {
		throw new AssistantToolError(raced)
	}

	useChatComposerStore.getState().setDraft(chat.uuid, text)

	return {
		value: { drafted: true, chat: name, note: "The message is in the input field; the user sends it." },
		uiEffect: { type: "drafted", pathname: `/chats/${chat.uuid}` }
	}
}
