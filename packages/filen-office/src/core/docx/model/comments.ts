import type { XmlDocument } from "../../xml/index"
import { Ns } from "./namespaces"
import { partReader, type PartReader } from "./reader"

export interface CommentInfo {
	readonly id: string
	readonly story: number
	readonly author: string | undefined
	readonly initials: string | undefined
	readonly date: string | undefined
	// w14:paraId of the comment's last paragraph, which commentsExtended and commentsIds key on.
	readonly paraId: string | undefined
	// The comment this one replies to (commentsExtended paraIdParent), by w:id.
	readonly parentId: string | undefined
	readonly done: boolean
	readonly durableId: string | undefined
	readonly dateUtc: string | undefined
	// Row in the model's comment anchors, or -1 when nothing in the document refers to it.
	readonly anchor: number
}

export interface Person {
	readonly author: string | undefined
	readonly providerId: string | undefined
	readonly userId: string | undefined
}

export interface CommentExtension {
	readonly parentParaId: string | undefined
	readonly done: boolean
}

function children(
	document: XmlDocument,
	rootGroup: Ns,
	rootName: string,
	childName: string,
	visit: (reader: PartReader, node: number) => void
): void {
	const reader = partReader(document)
	const root = document.documentElement()

	if (!reader.is(root, rootGroup, rootName)) return

	for (let child = document.firstChildElement(root); child >= 0; child = document.nextSiblingElement(child)) {
		if (reader.is(child, rootGroup, childName)) visit(reader, child)
	}
}

// word/commentsExtended.xml: reply threading and resolved state, keyed by paraId.
export function readCommentsExtended(document: XmlDocument): Map<string, CommentExtension> {
	const out = new Map<string, CommentExtension>()

	children(document, Ns.W15, "commentsEx", "commentEx", (reader, node) => {
		const paraId = reader.attr(node, Ns.W15, "paraId")

		if (paraId !== undefined) {
			out.set(paraId.toUpperCase(), {
				parentParaId: reader.attr(node, Ns.W15, "paraIdParent")?.toUpperCase(),
				done: reader.flag(node, Ns.W15, "done")
			})
		}
	})

	return out
}

// word/commentsIds.xml: paraId -> durableId.
export function readCommentsIds(document: XmlDocument): Map<string, string> {
	const out = new Map<string, string>()

	children(document, Ns.W16Cid, "commentsIds", "commentId", (reader, node) => {
		const paraId = reader.attr(node, Ns.W16Cid, "paraId")
		const durableId = reader.attr(node, Ns.W16Cid, "durableId")

		if (paraId !== undefined && durableId !== undefined) out.set(paraId.toUpperCase(), durableId.toUpperCase())
	})

	return out
}

// word/commentsExtensible.xml: durableId -> UTC date.
export function readCommentsExtensible(document: XmlDocument): Map<string, string> {
	const out = new Map<string, string>()

	children(document, Ns.W16Cex, "commentsExtensible", "commentExtensible", (reader, node) => {
		const durableId = reader.attr(node, Ns.W16Cex, "durableId")
		const date = reader.attr(node, Ns.W16Cex, "dateUtc")

		if (durableId !== undefined && date !== undefined) out.set(durableId.toUpperCase(), date)
	})

	return out
}

// word/people.xml.
export function readPeople(document: XmlDocument): Person[] {
	const out: Person[] = []

	children(document, Ns.W15, "people", "person", (reader, node) => {
		const presence = reader.child(node, Ns.W15, "presenceInfo")

		out.push({
			author: reader.attr(node, Ns.W15, "author"),
			providerId: presence < 0 ? undefined : reader.attr(presence, Ns.W15, "providerId"),
			userId: presence < 0 ? undefined : reader.attr(presence, Ns.W15, "userId")
		})
	})

	return out
}
