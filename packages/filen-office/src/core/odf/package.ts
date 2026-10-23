import {
	ZipArchive,
	ZipError,
	type ByteSource,
	type PartChange,
	type Slicing,
	type ZipEntry,
	type ZipOpenOptions,
	type ZipPlan
} from "../opc/index"
import { FLAG_ENCRYPTED, METHOD_STORED } from "../opc/zipReader"
import { decodeUtf8, encodeUtf8, parseXml, parseXmlSliced, XmlError, type XmlDocument } from "../xml/index"
import { OdfError, type OdfOpenFailure } from "./errors"
import { MANIFEST_PATH, parseManifest, type OdfManifest } from "./manifest"

export type OdfKind = "text" | "spreadsheet" | "presentation" | "graphics" | "chart" | "image" | "formula" | "text-master" | "text-web"

export interface OdfMediaType {
	readonly kind: OdfKind
	readonly template: boolean
}

const MEDIA_PREFIX = "application/vnd.oasis.opendocument."

// The document media types of ODF 1.3 Part 2 Appendix C (text-master-template is new in 1.3).
export function odfMediaType(mediaType: string): OdfMediaType | undefined {
	if (!mediaType.startsWith(MEDIA_PREFIX)) return undefined

	const rest = mediaType.slice(MEDIA_PREFIX.length)
	const template = rest.endsWith("-template")
	const kind = template ? rest.slice(0, -"-template".length) : rest

	switch (kind) {
		case "text":
		case "spreadsheet":
		case "presentation":
		case "graphics":
		case "chart":
		case "image":
		case "formula":
		case "text-master":
			return { kind, template }
		case "text-web":
			return template ? undefined : { kind, template }
		default:
			return undefined
	}
}

export const MIMETYPE_PATH = "mimetype"
// A mimetype stream longer than this is not one (media types are short).
const MAX_MIMETYPE = 256

// How the package's mimetype stream meets Part 2 §3.3: first entry, stored, no extra field, the same media type
// as the manifest's root entry. LibreOffice opens packages that miss any of these; they are recorded, not refused.
export interface OdfMimetypeFacts {
	readonly present: boolean
	readonly value: string | undefined
	readonly first: boolean
	readonly stored: boolean
	readonly extraField: boolean
}

export interface OdfEncryption {
	// Entries the manifest marks encrypted, with their algorithm URIs.
	readonly entries: readonly string[]
	readonly algorithms: readonly string[]
	readonly keyDerivations: readonly string[]
	readonly openPgp: boolean
	// Entries encrypted at the zip level (the general-purpose encryption flag), which no ODF producer writes.
	readonly zipEntries: readonly string[]
}

// Things a conforming reader tolerates; kept for diagnostics and for writers that fix them on save.
export type OdfPackageIssue =
	| { readonly kind: "mimetype"; readonly facts: OdfMimetypeFacts }
	| { readonly kind: "mediaTypeMismatch"; readonly mimetype: string; readonly manifest: string }
	| { readonly kind: "noManifest" }
	| { readonly kind: "missingEntry"; readonly path: string }
	| { readonly kind: "unlistedEntry"; readonly path: string }

export interface OdfFacts {
	readonly kind: OdfKind
	readonly template: boolean
	readonly mediaType: string
	// manifest:version of the package (absent before ODF 1.2).
	readonly version: string | undefined
	readonly mimetype: OdfMimetypeFacts
	// XML signatures over the package (META-INF/documentsignatures.xml): an edit invalidates them.
	readonly signed: boolean
	// Basic or script libraries in the package, kept and never run.
	readonly macros: boolean
	readonly issues: readonly OdfPackageIssue[]
}

export type OdfOpenResult =
	| { readonly ok: true; readonly package: OdfPackage }
	| { readonly ok: false; readonly reason: OdfOpenFailure; readonly message: string; readonly encryption?: OdfEncryption }

export interface OdfOpenOptions extends ZipOpenOptions {
	// Refuse a package of another kind with "wrongKind" (templates of the kind are accepted).
	readonly kinds?: readonly OdfKind[]
}

function zipFailure(error: ZipError): OdfOpenFailure {
	switch (error.code) {
		case "notZip":
			return "notZip"
		case "compoundFile":
			return "compoundFile"
		case "tooManyEntries":
		case "tooLarge":
			return "tooLarge"
		case "encrypted":
			return "encrypted"
		default:
			return "invalidPackage"
	}
}

function failure(error: unknown): OdfOpenResult {
	if (error instanceof OdfError) return { ok: false, reason: error.reason, message: error.message }
	if (error instanceof ZipError) return { ok: false, reason: zipFailure(error), message: error.message }
	if (error instanceof XmlError) return { ok: false, reason: "invalidPackage", message: error.message }

	throw error
}

function encryptionOf(archive: ZipArchive, manifest: OdfManifest | undefined): OdfEncryption | undefined {
	const entries: string[] = []
	const algorithms = new Set<string>()
	const keyDerivations = new Set<string>()

	for (const entry of manifest?.entries ?? []) {
		if (entry.encryption === undefined) continue

		entries.push(entry.fullPath)

		if (entry.encryption.algorithm !== undefined) algorithms.add(entry.encryption.algorithm)
		if (entry.encryption.keyDerivation !== undefined) keyDerivations.add(entry.encryption.keyDerivation)
	}

	const zipEntries = archive.entries.filter(entry => (entry.flags & FLAG_ENCRYPTED) !== 0).map(entry => entry.name)
	const openPgp = manifest?.keyInfo === true

	if (entries.length === 0 && zipEntries.length === 0 && !openPgp) return undefined

	return { entries, algorithms: [...algorithms], keyDerivations: [...keyDerivations], openPgp, zipEntries }
}

const BODY_STREAMS: Readonly<Record<OdfKind, readonly string[]>> = {
	text: ["content.xml"],
	spreadsheet: ["content.xml"],
	presentation: ["content.xml"],
	graphics: ["content.xml"],
	chart: ["content.xml"],
	formula: ["content.xml"],
	"text-master": ["content.xml"],
	"text-web": ["content.xml"],
	// An image document's content is its picture; content.xml is optional.
	image: []
}

// An opened ODF package: the zip, its manifest and the facts about it. Streams are read on request only.
export class OdfPackage {
	readonly archive: ZipArchive
	readonly manifest: OdfManifest | undefined
	readonly facts: OdfFacts

	private constructor(archive: ZipArchive, manifest: OdfManifest | undefined, facts: OdfFacts) {
		this.archive = archive
		this.manifest = manifest
		this.facts = facts
	}

	static async open(source: ByteSource, options: OdfOpenOptions = {}): Promise<OdfOpenResult> {
		try {
			return { ok: true, package: await OdfPackage.read(await ZipArchive.open(source, options), options) }
		} catch (error) {
			if (error instanceof EncryptedPackage) {
				return { ok: false, reason: "encrypted", message: error.message, encryption: error.encryption }
			}

			return failure(error)
		}
	}

	private static async read(archive: ZipArchive, options: OdfOpenOptions): Promise<OdfPackage> {
		const issues: OdfPackageIssue[] = []
		const mimetypeEntry = archive.get(MIMETYPE_PATH)
		let mimetypeValue: string | undefined

		if (mimetypeEntry !== undefined && mimetypeEntry.uncompressedSize <= MAX_MIMETYPE) {
			mimetypeValue = decodeUtf8(await archive.read(mimetypeEntry), 0, mimetypeEntry.uncompressedSize).trim()
		}

		const first = archive.entries.reduce<ZipEntry | undefined>(
			(lowest, entry) => (lowest === undefined || entry.localHeaderOffset < lowest.localHeaderOffset ? entry : lowest),
			undefined
		)
		const mimetype: OdfMimetypeFacts = {
			present: mimetypeEntry !== undefined,
			value: mimetypeValue,
			first: mimetypeEntry !== undefined && first === mimetypeEntry,
			stored: mimetypeEntry?.method === METHOD_STORED,
			extraField: mimetypeEntry !== undefined && (await localExtraLength(archive, mimetypeEntry)) > 0
		}

		if (mimetype.present && (!mimetype.first || !mimetype.stored || mimetype.extraField))
			issues.push({ kind: "mimetype", facts: mimetype })

		const manifestEntry = archive.get(MANIFEST_PATH)
		let manifest: OdfManifest | undefined

		if (manifestEntry === undefined) {
			issues.push({ kind: "noManifest" })
		} else {
			manifest = parseManifest(await archive.read(manifestEntry))
		}

		const rootType = manifest?.entries.find(entry => entry.fullPath === "/")?.mediaType
		const mediaType = mimetypeValue ?? rootType

		if (mediaType === undefined || odfMediaType(mediaType) === undefined) {
			if (archive.get("content.xml") !== undefined && (await isLegacyOffice(archive))) {
				throw new OdfError("notOdf", "A StarOffice/OpenOffice.org 1.x file, not OpenDocument")
			}

			throw new OdfError("notOdf", "No OpenDocument media type in mimetype or the manifest")
		}

		if (mimetypeValue !== undefined && rootType !== undefined && rootType !== mimetypeValue) {
			issues.push({ kind: "mediaTypeMismatch", mimetype: mimetypeValue, manifest: rootType })
		}

		const type = odfMediaType(mediaType) ?? { kind: "text", template: false }
		const encryption = encryptionOf(archive, manifest)

		if (encryption !== undefined) throw new EncryptedPackage(encryption)

		if (options.kinds !== undefined && !options.kinds.includes(type.kind)) {
			throw new OdfError("wrongKind", `An OpenDocument ${type.kind} document`)
		}

		for (const stream of BODY_STREAMS[type.kind]) {
			if (archive.get(stream) === undefined) throw new OdfError("invalidPackage", `The package has no ${stream}`)
		}

		if (manifest !== undefined) {
			const listed = new Set<string>()

			for (const entry of manifest.entries) {
				if (entry.fullPath === "/" || entry.fullPath.endsWith("/")) continue

				listed.add(entry.fullPath)

				if (archive.get(entry.fullPath) === undefined) issues.push({ kind: "missingEntry", path: entry.fullPath })
			}

			for (const entry of archive.entries) {
				if (entry.name.endsWith("/") || entry.name === MIMETYPE_PATH || entry.name.startsWith("META-INF/")) continue
				if (!listed.has(entry.name)) issues.push({ kind: "unlistedEntry", path: entry.name })
			}
		}

		const signed = archive.get("META-INF/documentsignatures.xml") !== undefined
		const macros = archive.entries.some(entry => entry.name.startsWith("Basic/") || entry.name.startsWith("Scripts/"))

		return new OdfPackage(archive, manifest, {
			kind: type.kind,
			template: type.template,
			mediaType,
			version: manifest?.version,
			mimetype,
			signed,
			macros,
			issues
		})
	}

	has(path: string): boolean {
		return this.archive.get(path) !== undefined
	}

	// A stream's bytes, or undefined when the package has none by that path.
	async bytes(path: string, slicing?: Slicing): Promise<Uint8Array | undefined> {
		const entry = this.archive.get(path)

		if (entry === undefined) return undefined

		return slicing === undefined ? this.archive.read(entry) : this.archive.readSliced(entry, slicing)
	}

	async xml(path: string, slicing?: Slicing): Promise<XmlDocument | undefined> {
		const bytes = await this.bytes(path, slicing)

		if (bytes === undefined) return undefined

		return slicing === undefined ? parseXml(bytes) : parseXmlSliced(bytes, slicing)
	}

	// The media type of a stream: the manifest's, else undefined.
	mediaTypeOf(path: string): string | undefined {
		return this.manifest?.entries.find(entry => entry.fullPath === path)?.mediaType
	}
}

class EncryptedPackage extends Error {
	readonly encryption: OdfEncryption

	constructor(encryption: OdfEncryption) {
		super(
			encryption.openPgp
				? "The package is encrypted with OpenPGP keys"
				: `The package is encrypted (${encryption.algorithms.join(", ") || "zip encryption"})`
		)

		this.name = "EncryptedPackage"
		this.encryption = encryption
	}
}

async function localExtraLength(archive: ZipArchive, entry: ZipEntry): Promise<number> {
	const header = await archive.source.read(entry.localHeaderOffset, 30)

	return header.length < 30 ? 0 : (header[28] ?? 0) | ((header[29] ?? 0) << 8)
}

const LEGACY_SNIFF_LIMIT = 4 * 1024 * 1024

async function isLegacyOffice(archive: ZipArchive): Promise<boolean> {
	const entry = archive.get("content.xml")

	// Only for telling apart a refusal's message: a large stream is not inflated for it.
	if (entry === undefined || entry.uncompressedSize > LEGACY_SNIFF_LIMIT) return false

	const head = await archive.read(entry)

	return decodeUtf8(head, 0, Math.min(head.length, 2048)).includes("http://openoffice.org/2000/office")
}

export interface NewOdfPackage {
	readonly mediaType: string
	// Every other stream, META-INF/manifest.xml included, in the order they are written.
	readonly parts: readonly PartChange[]
}

// A new package's zip plan: the mimetype stream first, stored, without an extra field (Part 2 §3.3), so its
// media type sits at byte 38 where content sniffers look for it.
export function newOdfPlan(init: NewOdfPackage): ZipPlan {
	return {
		items: [
			{ kind: "data", name: MIMETYPE_PATH, bytes: encodeUtf8(init.mediaType), compress: false },
			...init.parts.map(part => ({ kind: "data" as const, name: part.name, bytes: part.bytes, compress: part.compress ?? true }))
		]
	}
}
