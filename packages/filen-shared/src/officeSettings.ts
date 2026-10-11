// The user's document-editing settings as stored (end-to-end encrypted, as a file the apps sync next to the
// user dictionary): the author name comments and tracked changes carry, and whether edits save on their own.
// Each setting is a last-writer-wins register, so devices that change it offline merge without one undoing
// the other unseen. An unset setting has no register (the author name: never asked yet).
//
// { "v": 1, "authorName": { "value": "<name>", "at": <stamp> }, "autosave": { "value": <boolean>, "at": <stamp> } }
//
// Stamps are milliseconds since the epoch; on a tie the larger value (in JSON text order) wins, so every device
// merges to the same document.

export const OFFICE_SETTINGS_FILE_NAME = "settings.v1.json"
export const OFFICE_SETTINGS_VERSION = 1
// Word's own limit for a user name.
export const MAX_AUTHOR_NAME_LENGTH = 255
export const DEFAULT_OFFICE_AUTOSAVE = true

export interface OfficeSetting<T> {
	readonly value: T
	readonly at: number
}

export interface OfficeSettingsDoc {
	readonly v: typeof OFFICE_SETTINGS_VERSION
	readonly authorName?: OfficeSetting<string>
	readonly autosave?: OfficeSetting<boolean>
}

// The file was written by a newer format: it must not be overwritten by this one.
export class OfficeSettingsVersionError extends Error {
	readonly version: unknown

	constructor(version: unknown) {
		super(`Unsupported office settings version ${String(version)}`)
		this.name = "OfficeSettingsVersionError"
		this.version = version
	}
}

export function emptyOfficeSettings(): OfficeSettingsDoc {
	return { v: OFFICE_SETTINGS_VERSION }
}

// The form an author name is stored in (NFC, inner whitespace collapsed, trimmed), or undefined when it is
// empty or too long.
export function normalizeAuthorName(name: string): string | undefined {
	const normalized = name.normalize("NFC").replace(/\s+/gu, " ").trim()

	// Control characters (other than the whitespace collapsed above) never belong in a name.
	if (normalized.length === 0 || normalized.length > MAX_AUTHOR_NAME_LENGTH || /\p{Cc}/u.test(normalized)) return undefined

	return normalized
}

function clampStamp(stamp: number): number {
	return Math.min(Math.max(Math.trunc(stamp), -Number.MAX_SAFE_INTEGER), Number.MAX_SAFE_INTEGER)
}

function register<T>(value: unknown, read: (raw: unknown) => T | undefined): OfficeSetting<T> | undefined {
	if (typeof value !== "object" || value === null || Array.isArray(value)) return undefined

	const at: unknown = Reflect.get(value, "at")
	const parsed = read(Reflect.get(value, "value"))

	if (parsed === undefined || typeof at !== "number" || !Number.isFinite(at)) return undefined

	return { value: parsed, at: clampStamp(at) }
}

function readName(raw: unknown): string | undefined {
	return typeof raw === "string" ? normalizeAuthorName(raw) : undefined
}

function readBoolean(raw: unknown): boolean | undefined {
	return typeof raw === "boolean" ? raw : undefined
}

function build(authorName: OfficeSetting<string> | undefined, autosave: OfficeSetting<boolean> | undefined): OfficeSettingsDoc {
	return {
		v: OFFICE_SETTINGS_VERSION,
		...(authorName === undefined ? {} : { authorName }),
		...(autosave === undefined ? {} : { autosave })
	}
}

// Reads a stored file; settings that are not valid are dropped. Throws OfficeSettingsVersionError for a newer
// format and SyntaxError for text that is not JSON.
export function parseOfficeSettings(text: string): OfficeSettingsDoc {
	const value: unknown = JSON.parse(text)

	if (typeof value !== "object" || value === null || Array.isArray(value)) return emptyOfficeSettings()

	const version: unknown = Reflect.get(value, "v")

	if (typeof version === "number" && version > OFFICE_SETTINGS_VERSION) throw new OfficeSettingsVersionError(version)

	return build(register(Reflect.get(value, "authorName"), readName), register(Reflect.get(value, "autosave"), readBoolean))
}

// Stable text (fixed key order), so equal documents serialise identically.
export function serializeOfficeSettings(doc: OfficeSettingsDoc): string {
	const pick = <T>(setting: OfficeSetting<T> | undefined) =>
		setting === undefined ? undefined : { value: setting.value, at: setting.at }

	return JSON.stringify({ v: OFFICE_SETTINGS_VERSION, authorName: pick(doc.authorName), autosave: pick(doc.autosave) })
}

function later<T>(a: OfficeSetting<T> | undefined, b: OfficeSetting<T> | undefined): OfficeSetting<T> | undefined {
	if (a === undefined) return b
	if (b === undefined || a.at > b.at) return a
	if (b.at > a.at) return b

	return JSON.stringify(a.value) >= JSON.stringify(b.value) ? a : b
}

export function mergeOfficeSettings(a: OfficeSettingsDoc, b: OfficeSettingsDoc): OfficeSettingsDoc {
	return build(later(a.authorName, b.authorName), later(a.autosave, b.autosave))
}

// Whether two documents hold the same settings (an unchanged merge needs no upload).
export function sameOfficeSettings(a: OfficeSettingsDoc, b: OfficeSettingsDoc): boolean {
	return serializeOfficeSettings(a) === serializeOfficeSettings(b)
}

// The document with the author name set at `stamp` (unchanged when the name is not valid or already set so).
export function setAuthorName(doc: OfficeSettingsDoc, name: string, stamp: number): OfficeSettingsDoc {
	const normalized = normalizeAuthorName(name)

	if (normalized === undefined || doc.authorName?.value === normalized) return doc

	// Past the current register even when this device's clock is behind the one that set it.
	return build({ value: normalized, at: clampStamp(Math.max(stamp, (doc.authorName?.at ?? Number.NEGATIVE_INFINITY) + 1)) }, doc.autosave)
}

export function setOfficeAutosave(doc: OfficeSettingsDoc, autosave: boolean, stamp: number): OfficeSettingsDoc {
	if (officeAutosave(doc) === autosave && doc.autosave !== undefined) return doc

	return build(doc.authorName, { value: autosave, at: clampStamp(Math.max(stamp, (doc.autosave?.at ?? Number.NEGATIVE_INFINITY) + 1)) })
}

export function officeAuthorName(doc: OfficeSettingsDoc): string | null {
	return doc.authorName?.value ?? null
}

export function officeAutosave(doc: OfficeSettingsDoc): boolean {
	return doc.autosave?.value ?? DEFAULT_OFFICE_AUTOSAVE
}
