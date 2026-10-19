import type { Transaction } from "prosemirror-state"
import type { StoryKind } from "@filen/office/core/docx"
import { nameHash, rng, type Rand } from "../config"
import type { DocxEditorCommand, DocxEditorSession, PagedEditor } from "../harness"
import { baseOps } from "./base"
import { commentOps } from "./comments"
import { drawingOps } from "./drawings"
import { editorOps } from "./editor"
import { fieldOps } from "./fields"
import { formattingOps } from "./formatting"
import { formOps } from "./forms"
import { headerOps } from "./headers"
import { mathOps } from "./math"
import { noteOps } from "./notes"
import { reviewOps } from "./review"
import { tableOps } from "./tables"
import { textboxOps } from "./textbox"

export type { Rand }

// The edit operations of the fuzzers and the scripted edit suite, one file per feature.

export interface OpContext {
	readonly rand: Rand
	// The session of the story being edited: under the paged fuzz, the paged editor's active story.
	readonly session: DocxEditorSession
	// Paged fuzz only.
	readonly editor: PagedEditor | undefined
	// Ids the source's numbering and styles parts define.
	readonly numIds: readonly number[]
	readonly styleIds: readonly string[]
	// Dispatches the transaction `build` makes and returns the outcome for the log ("true", "false", or why
	// it never applied); the damage fuzz records builds and dispatches that throw.
	transact(label: string, build: (tr: Transaction) => Transaction | void): string
	// Executes a command and returns the outcome for the log; the damage fuzz records one that throws.
	execute(label: string, command: DocxEditorCommand): string
}

export interface EditOp {
	readonly name: string
	// "paged": user-level edits under the paged editor, whose layout must equal a fresh layout of the saved
	// file. "damage": raw transactions aimed at the writer and Gate 2, read back by the independent reader.
	readonly harness: "paged" | "damage"
	// Relative pick weight within its feature (1 when absent; 0 runs only when named in the op filter).
	readonly weight?: number
	// Applies the op; returns a log entry ("noop" when nothing applied).
	run(context: OpContext): string | Promise<string>
	// The deterministic edit the scripted suite makes, if any.
	readonly script?: ScriptEdit
	// Parts the scripted edit may write besides the main part, styles, numbering and relationships, by content
	// type: a full type, or the name before "+xml" ("comments" for ...wordprocessingml.comments+xml).
	readonly writes?: readonly string[]
}

export interface ScriptEdit {
	// Place in the scripted suite (ascending), independent of the registry order: a later edit sees the
	// earlier ones, so existing orders never change.
	readonly order: number
	// True when it applied.
	run(target: ScriptTarget): boolean
}

export interface ScriptTarget {
	// The body's session.
	readonly session: DocxEditorSession
	// A session for the document's first story of `kind` (the same one on every call), if it has one.
	story(kind: StoryKind): DocxEditorSession | undefined
}

interface Feature {
	readonly name: string
	readonly ops: readonly EditOp[]
	// Fixed share of each harness's steps. A feature without ops for the harness lends its share to base,
	// drawing on a stream of its own, so filling it in later changes no other feature's draws.
	readonly paged: number
	readonly damage: number
}

const FEATURES: readonly Feature[] = [
	{ name: "base", ops: baseOps, paged: 12, damage: 12 },
	{ name: "editor", ops: editorOps, paged: 3, damage: 0 },
	{ name: "comments", ops: commentOps, paged: 1, damage: 1 },
	{ name: "review", ops: reviewOps, paged: 1, damage: 1 },
	{ name: "tables", ops: tableOps, paged: 1, damage: 1 },
	{ name: "formatting", ops: formattingOps, paged: 1, damage: 1 },
	{ name: "notes", ops: noteOps, paged: 1, damage: 1 },
	{ name: "headers", ops: headerOps, paged: 1, damage: 1 },
	{ name: "fields", ops: fieldOps, paged: 1, damage: 1 },
	{ name: "forms", ops: formOps, paged: 1, damage: 1 },
	{ name: "drawings", ops: drawingOps, paged: 1, damage: 1 },
	{ name: "textbox", ops: textboxOps, paged: 1, damage: 1 },
	{ name: "math", ops: mathOps, paged: 1, damage: 1 }
]

export const EDIT_OPS: readonly EditOp[] = FEATURES.flatMap(feature => feature.ops)

// The features (op families) with ops for `harness`.
export function featureNames(harness: EditOp["harness"]): string[] {
	return FEATURES.filter(feature => feature[harness] > 0 && feature.ops.some(op => op.harness === harness && (op.weight ?? 1) > 0)).map(
		feature => feature.name
	)
}

// Picks one op by weight with one draw.
function pickWeighted(ops: readonly EditOp[], rand: Rand): EditOp {
	const weight = (op: EditOp): number => Math.max(op.weight ?? 1, 1)
	let at = Math.floor(rand() * ops.reduce((sum, op) => sum + weight(op), 0))

	for (const op of ops) {
		at -= weight(op)

		if (at < 0) return op
	}

	throw new Error("No ops to pick from")
}

function stream(seed: number, name: string): Rand {
	return rng((Math.imul(seed, 0x9e3779b1) ^ nameHash(name)) >>> 0)
}

// Chooses each step's op. The run's own stream picks the feature by the fixed shares above; the op, and every
// draw the op makes, come from that feature's own stream (a base op on a lent share from the lending feature's
// own lent stream). A seed therefore replays every step up to the first one that picks a feature whose ops
// changed, and later steps still pick the same features.
// With `names`, it picks uniformly among those ops on the run's own stream instead. With `families`, only those
// features take shares, and none lends to base.
export class OpPicker {
	private readonly features: {
		readonly feature: Feature
		readonly ops: EditOp[]
		readonly rand: Rand
		readonly lent: Rand
		readonly share: number
	}[]
	private readonly total: number
	private readonly named: EditOp[] | undefined

	constructor(harness: EditOp["harness"], seed: number, names?: readonly string[], families?: readonly string[]) {
		const unknown = families?.filter(family => !featureNames(harness).includes(family)) ?? []

		if (unknown.length > 0) throw new Error(`No ${harness} ops in ${unknown.join(", ")}`)

		this.features = FEATURES.filter(feature => families === undefined || families.includes(feature.name))
			.map(feature => ({
				feature,
				ops: feature.ops.filter(op => op.harness === harness && (op.weight ?? 1) > 0),
				rand: stream(seed, feature.name),
				lent: stream(seed, `${feature.name}:lent`),
				share: feature[harness]
			}))
			.filter(entry => entry.share > 0)
		this.total = this.features.reduce((sum, entry) => sum + entry.share, 0)
		this.named = names === undefined ? undefined : EDIT_OPS.filter(op => op.harness === harness && names.includes(op.name))

		if (this.named?.length === 0) throw new Error(`No ${harness} ops named ${names?.join(", ")}`)
	}

	next(run: Rand): { readonly op: EditOp; readonly rand: Rand } {
		if (this.named !== undefined) return { op: pickWeighted(this.named, run), rand: run }

		let at = Math.floor(run() * this.total)
		const base = (this.features.find(entry => entry.feature.name === "base") ?? this.features[0]) as { readonly ops: EditOp[] }

		for (const entry of this.features) {
			at -= entry.share

			if (at >= 0) continue
			if (entry.ops.length > 0) return { op: pickWeighted(entry.ops, entry.rand), rand: entry.rand }

			return { op: pickWeighted(base.ops, entry.lent), rand: entry.lent }
		}

		throw new Error("No feature to pick from")
	}
}

// The deterministic subset, in script order.
export function scriptOps(): EditOp[] {
	return EDIT_OPS.filter(op => op.script !== undefined).sort((a, b) => (a.script?.order ?? 0) - (b.script?.order ?? 0))
}
