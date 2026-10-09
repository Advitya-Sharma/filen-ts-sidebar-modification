// The synced Videos library configuration, `.filen/Videos/library.v1.json`: which directories are library roots
// (and how each is classified), library preferences, and manual corrections. Written rarely, by any device.
//
// Built from the watch state's register records: each root, preference and correction is a record of
// last-writer-wins registers stamped with a hybrid logical clock and a writer id, and removing one writes a
// tombstone, so concurrent edits of different fields both survive and a merge with an older copy never brings a
// removed entry back. Records, registers and fields this version doesn't know (a root type added later, say) are
// carried through verbatim, so an older client rewriting the file never drops or rewrites what a newer one wrote.

import type { LibraryRootType } from "./classify"
import {
	hasLiveRegister,
	mergeRecordMaps,
	mergeUnknownFields,
	parseRegisterRecord,
	setRegister,
	tombstoneRecord,
	type Hlc,
	type RegisterRecord,
	type Stamp
} from "./watchState"

export const VIDEOS_DIR_NAME = "Videos"
export const LIBRARY_CONFIG_FILE_NAME = "library.v1.json"

const ROOT_TYPES: ReadonlySet<unknown> = new Set<LibraryRootType>(["auto", "movies", "shows", "other"])

export interface ValueRegister<V> extends Stamp {
	v: V
	[key: string]: unknown
}

// ty: root type, en: enabled, ad: wall-clock ms the root was added (orders the roots).
export interface LibraryRootRecord extends RegisterRecord {
	ty?: ValueRegister<unknown>
	en?: ValueRegister<unknown>
	ad?: ValueRegister<unknown>
}

export interface LibraryPrefRecord extends RegisterRecord {
	p?: ValueRegister<unknown>
}

// A manual correction, keyed `d:<directory uuid>` (a title with its own directory, a series) or
// `s:<stable uuid>` (a single file); one register, edited as a whole.
export interface LibraryOverride {
	title?: string
	year?: number
	kind?: "movie" | "series" | "other"
	hidden?: boolean
	// Key of the title this one is shown as part of.
	mergeInto?: string
}

export interface LibraryOverrideRecord extends RegisterRecord {
	o?: Stamp & LibraryOverride & Record<string, unknown>
}

export interface LibraryConfigDoc {
	v: 1
	roots: Record<string, LibraryRootRecord>
	prefs: Record<string, LibraryPrefRecord>
	overrides: Record<string, LibraryOverrideRecord>
	[key: string]: unknown
}

const DOC_FIELDS: ReadonlySet<string> = new Set(["v", "roots", "prefs", "overrides"])

export function emptyLibraryConfig(): LibraryConfigDoc {
	return {
		v: 1,
		roots: {},
		prefs: {},
		overrides: {}
	}
}

function isObject(value: unknown): value is Record<string, unknown> {
	return typeof value === "object" && value !== null && !Array.isArray(value)
}

function parseSection<T extends RegisterRecord>(value: unknown): Record<string, T> {
	const out: Record<string, T> = {}

	if (!isObject(value)) {
		return out
	}

	for (const key in value) {
		const record = key === "__proto__" ? null : parseRegisterRecord(value[key])

		if (record !== null) {
			out[key] = record as T
		}
	}

	return out
}

// null when the value isn't a version-1 library document at all.
export function parseLibraryConfig(value: unknown): LibraryConfigDoc | null {
	if (!isObject(value) || value["v"] !== 1) {
		return null
	}

	return {
		...value,
		v: 1,
		roots: parseSection(value["roots"]),
		prefs: parseSection(value["prefs"]),
		overrides: parseSection(value["overrides"])
	}
}

export function parseLibraryConfigText(text: string): LibraryConfigDoc | null {
	try {
		return parseLibraryConfig(JSON.parse(text) as unknown)
	} catch {
		return null
	}
}

export interface LibraryConfigMerge {
	doc: LibraryConfigDoc
	// local holds something remote lacks: an upload is due
	localContributed: boolean
	// remote holds something local lacks
	remoteContributed: boolean
}

// Register-wise merge; associative, commutative and idempotent.
export function mergeLibraryConfig({ local, remote }: { local: LibraryConfigDoc; remote: LibraryConfigDoc }): LibraryConfigMerge {
	const roots = mergeRecordMaps({ local: local.roots, remote: remote.roots })
	const prefs = mergeRecordMaps({ local: local.prefs, remote: remote.prefs })
	const overrides = mergeRecordMaps({ local: local.overrides, remote: remote.overrides })
	const fields = mergeUnknownFields({ local, remote, known: DOC_FIELDS })
	const parts = [roots, prefs, overrides, fields]

	return {
		doc: {
			...fields.value,
			v: 1,
			roots: roots.value,
			prefs: prefs.value,
			overrides: overrides.value
		},
		localContributed: parts.some(part => part.localContributed),
		remoteContributed: parts.some(part => part.remoteContributed)
	}
}

// The newest stamp in the document, for the clock to observe after a read.
export function libraryConfigMaxStamp(doc: LibraryConfigDoc): Hlc {
	let max = 0

	for (const section of [doc.roots, doc.prefs, doc.overrides]) {
		for (const key in section) {
			const record: RegisterRecord | undefined = section[key]

			for (const field in record) {
				const value = record[field]

				if (isObject(value) && typeof value["t"] === "number") {
					max = Math.max(max, value["t"])
				}
			}
		}
	}

	return max
}

function canonical(value: unknown): unknown {
	if (Array.isArray(value)) {
		return value.map(canonical)
	}

	if (!isObject(value)) {
		return value
	}

	const out: Record<string, unknown> = {}

	for (const key of Object.keys(value).sort()) {
		out[key] = canonical(value[key])
	}

	return out
}

// Key-order independent, so two documents with equal content serialize equally.
export function serializeLibraryConfig(doc: LibraryConfigDoc): string {
	return JSON.stringify(canonical(doc))
}

export interface ActiveLibraryRoot {
	dirUuid: string
	type: LibraryRootType
	enabled: boolean
	added: number
}

// The roots that aren't removed, in the order they were added. A type this version doesn't know reads as auto.
export function activeLibraryRoots(doc: LibraryConfigDoc): ActiveLibraryRoot[] {
	const out: ActiveLibraryRoot[] = []

	for (const dirUuid in doc.roots) {
		const record = doc.roots[dirUuid]

		if (record?.ty === undefined) {
			continue
		}

		const type = record.ty.v
		const added = record.ad?.v

		out.push({
			dirUuid,
			type: ROOT_TYPES.has(type) ? (type as LibraryRootType) : "auto",
			enabled: record.en?.v !== false,
			added: typeof added === "number" ? added : 0
		})
	}

	return out.sort((a, b) => a.added - b.added || (a.dirUuid < b.dirUuid ? -1 : a.dirUuid > b.dirUuid ? 1 : 0))
}

export function libraryPref(doc: LibraryConfigDoc, key: string): unknown {
	return doc.prefs[key]?.p?.v
}

export function libraryOverride(doc: LibraryConfigDoc, key: string): LibraryOverride | null {
	return doc.overrides[key]?.o ?? null
}

// ── Edits ────────────────────────────────────────────────────────────────────

function withRecord<K extends "roots" | "prefs" | "overrides">(
	doc: LibraryConfigDoc,
	section: K,
	key: string,
	record: LibraryConfigDoc[K][string]
): LibraryConfigDoc {
	return {
		...doc,
		[section]: {
			...doc[section],
			[key]: record
		}
	}
}

function register<V>(stamp: Stamp, value: V): ValueRegister<V> {
	return {
		t: stamp.t,
		d: stamp.d,
		v: value
	}
}

// Adds a root (or brings a removed one back), or changes its type or enabled flag. Only the registers given and
// changed are written, so another device's concurrent edit of the other one survives, and a value this version
// doesn't know is left alone unless the edit replaces it. A new root defaults to auto, enabled.
export function withLibraryRoot({
	doc,
	dirUuid,
	type,
	enabled,
	stamp,
	now
}: {
	doc: LibraryConfigDoc
	dirUuid: string
	type?: LibraryRootType
	enabled?: boolean
	stamp: Stamp
	now: number
}): LibraryConfigDoc {
	const previous = doc.roots[dirUuid]
	const live = previous?.ty !== undefined
	let record = previous

	if (!live || (type !== undefined && previous.ty?.v !== type)) {
		record = setRegister({ rec: record, key: "ty", register: register(stamp, type ?? "auto") })
	}

	if (!live || (enabled !== undefined && previous.en?.v !== enabled)) {
		record = setRegister({ rec: record, key: "en", register: register(stamp, enabled ?? true) })
	}

	if (!live || previous.ad === undefined) {
		record = setRegister({ rec: record, key: "ad", register: register(stamp, now) })
	}

	return record === previous || record === undefined ? doc : withRecord(doc, "roots", dirUuid, record)
}

export function withoutLibraryRoot({
	doc,
	dirUuid,
	stamp,
	at
}: {
	doc: LibraryConfigDoc
	dirUuid: string
	stamp: Stamp
	at: number
}): LibraryConfigDoc {
	const record = doc.roots[dirUuid]

	if (record === undefined || !hasLiveRegister(record)) {
		return doc
	}

	return withRecord(doc, "roots", dirUuid, tombstoneRecord({ rec: record, stamp, at }))
}

// Sets (or, with `value` undefined, removes) a library preference.
export function withLibraryPref({
	doc,
	key,
	value,
	stamp,
	at
}: {
	doc: LibraryConfigDoc
	key: string
	value: unknown
	stamp: Stamp
	at: number
}): LibraryConfigDoc {
	const record = doc.prefs[key]

	return withRecord(
		doc,
		"prefs",
		key,
		value === undefined
			? tombstoneRecord({ rec: record, stamp, at })
			: setRegister({ rec: record, key: "p", register: register(stamp, value) })
	)
}

// Sets (or, with `override` null, removes) the correction stored under `key`.
export function withLibraryOverride({
	doc,
	key,
	override,
	stamp,
	at
}: {
	doc: LibraryConfigDoc
	key: string
	override: LibraryOverride | null
	stamp: Stamp
	at: number
}): LibraryConfigDoc {
	const record = doc.overrides[key]

	return withRecord(
		doc,
		"overrides",
		key,
		override === null
			? tombstoneRecord({ rec: record, stamp, at })
			: setRegister({
					rec: record,
					key: "o",
					register: {
						...override,
						t: stamp.t,
						d: stamp.d
					}
				})
	)
}

// ── Root validation ──────────────────────────────────────────────────────────

// A directory as the root chooser sees it: its uuid, and the uuids from the drive root (inclusive) down to its
// parent; empty for the drive root itself. `ownDrive` is false for shared-in and public-link directories, which have
// no stable file ids.
export interface LibraryRootCandidate {
	uuid: string
	ancestors: readonly string[]
	ownDrive: boolean
}

export type LibraryRootValidation =
	| { ok: true }
	| { ok: false; reason: "notOwnDrive" | "systemDirectory" | "alreadyRoot" }
	// Overlaps existing roots (inside one, or holding some): offered as a replacement for them.
	| { ok: false; reason: "nested"; replace: string[] }

export function validateLibraryRoot({
	candidate,
	roots,
	dotFilenUuid
}: {
	candidate: LibraryRootCandidate
	// The current roots, each with its own ancestor chain.
	roots: readonly { uuid: string; ancestors: readonly string[] }[]
	// The `.filen` directory's uuid when known.
	dotFilenUuid: string | null
}): LibraryRootValidation {
	if (!candidate.ownDrive) {
		return {
			ok: false,
			reason: "notOwnDrive"
		}
	}

	if (dotFilenUuid !== null && (candidate.uuid === dotFilenUuid || candidate.ancestors.includes(dotFilenUuid))) {
		return {
			ok: false,
			reason: "systemDirectory"
		}
	}

	const replace: string[] = []

	for (const root of roots) {
		if (root.uuid === candidate.uuid) {
			return {
				ok: false,
				reason: "alreadyRoot"
			}
		}

		if (candidate.ancestors.includes(root.uuid) || root.ancestors.includes(candidate.uuid)) {
			replace.push(root.uuid)
		}
	}

	return replace.length > 0
		? {
				ok: false,
				reason: "nested",
				replace
			}
		: {
				ok: true
			}
}
