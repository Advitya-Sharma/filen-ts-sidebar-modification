import { Plugin, PluginKey, type EditorState, type Transaction } from "prosemirror-state"
import { codePointPlugin } from "./codePoints"
import { docxInvariants, isInternal, type InvariantOptions } from "./core"
import { FIELDS_STAGE } from "./fields"
import { FORMS_STAGE } from "./forms"
import { TABLES_STAGE } from "./tables"
import { TRACKING_STAGE } from "./tracking"

export {
	COMMENT_MOVE_META,
	changedRanges,
	docxInvariants,
	docxInvariantsKey,
	freshSlice,
	freshSliceKeeping,
	invariantViolation,
	isCommand,
	isInternal,
	markCommand,
	markInternal,
	normalizeListAttrs,
	repairTransaction,
	writableNew,
	type InvariantOptions,
	type ListAttrs,
	type ListNormalizer
} from "./core"

// The invariants as a pipeline: the core rules (structure, identity, markers, the read-only gate), then
// each feature's own. A stage refuses transactions that break what it keeps, and repairs after the
// transactions applied (it sees every transaction, the user's and the repairs of stages before it).
export interface InvariantStage {
	readonly name: string
	// Why `tr` would break the feature's structure, or undefined. The session's own transactions never come.
	readonly violation?: (tr: Transaction) => string | undefined
	// What to append after `transactions`, or null.
	readonly repair?: (transactions: readonly Transaction[], state: EditorState, options: InvariantOptions) => Transaction | null
}

export const INVARIANT_STAGES: readonly InvariantStage[] = [TRACKING_STAGE, FORMS_STAGE, FIELDS_STAGE, TABLES_STAGE].filter(
	(stage): stage is InvariantStage => stage !== null
)

// The plugins of an editing session: the core first, then whole code points, a plugin per feature stage after it.
export function docxInvariantPlugins(options: InvariantOptions = {}): Plugin[] {
	return [
		docxInvariants(options),
		codePointPlugin(),
		...INVARIANT_STAGES.map(
			stage =>
				new Plugin({
					key: new PluginKey(`docxInvariants.${stage.name}`),
					...(stage.violation === undefined
						? {}
						: {
								filterTransaction: (tr: Transaction) =>
									!tr.docChanged || isInternal(tr) || (stage.violation?.(tr) ?? undefined) === undefined
							}),
					...(stage.repair === undefined
						? {}
						: {
								appendTransaction: (transactions: readonly Transaction[], _old: EditorState, state: EditorState) =>
									stage.repair?.(transactions, state, options) ?? null
							})
				})
		)
	]
}
