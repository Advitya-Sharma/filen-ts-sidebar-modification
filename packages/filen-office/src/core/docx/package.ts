import {
	OpcError,
	OpcPackage,
	ZipError,
	type ByteSource,
	type ContentTypes,
	type Relationships,
	type Slicing,
	type ZipOpenOptions
} from "../opc/index"
import { XmlError, type XmlDocument } from "../xml/index"
import { NO_SETTINGS, readSettings, type DocumentProtection, type DocxSettings, type WriteProtection } from "./model/settings"

export type DocxVariant = "docx" | "docm" | "dotx" | "dotm"

// Why a file could not be opened as a word-processing package at all.
export type DocxOpenFailure =
	// OOXML encrypted with a password (an OLE2 compound file holding EncryptionInfo + EncryptedPackage).
	| "encrypted"
	// A legacy binary Office file (.doc, .xls, .ppt) under a .docx name.
	| "legacyBinary"
	// Some other OLE2 compound file.
	| "compoundFile"
	| "notZip"
	| "invalidPackage"
	| "tooLarge"
	// A valid package that is not a word-processing document (a workbook, a presentation...).
	| "notWordprocessing"

export interface DocxFacts {
	readonly variant: DocxVariant
	// ISO 29500 Strict namespaces (relationship types or settings).
	readonly strict: boolean
	readonly macros: boolean
	readonly template: boolean
	// Signatures in the package's digital signature origin.
	readonly signatures: number
	readonly documentProtection: DocumentProtection | undefined
	readonly writeProtection: WriteProtection | undefined
	readonly glossary: boolean
	// A template or sub-documents referenced from outside the package: recorded, never loaded.
	readonly attachedTemplate: boolean
	readonly subDocuments: number
	// Settings or their relationships could not be parsed (read as absent).
	readonly malformed: boolean
}

// Reasons a document opens view-only (gate 0 of the save proof). altChunk is not one of them: it is
// preserved and never rendered; nor is a template (.dotx/.dotm), saved with the content type it has, or a
// macro project, copied byte for byte and never run. "malformed" also covers parts read later, passed after
// the facts.
export type DocxGateReason = "strict" | "signed" | "protected" | "writeProtected" | "malformed"

// `parts`: anything else read for the document that knows whether its markup was malformed (the model,
// styles, numbering, theme).
export function gateReasons(facts: DocxFacts, ...parts: readonly ({ readonly malformed: boolean } | undefined)[]): DocxGateReason[] {
	const reasons: DocxGateReason[] = []
	const protection = facts.documentProtection

	if (facts.strict) reasons.push("strict")
	if (facts.signatures > 0) reasons.push("signed")
	if (protection !== undefined && protection.enforced && protection.edit !== undefined && protection.edit !== "none")
		reasons.push("protected")
	if (facts.writeProtection?.hashed === true) reasons.push("writeProtected")
	if (facts.malformed || parts.some(part => part?.malformed === true)) reasons.push("malformed")

	return reasons
}

const RELATIONSHIP_BASES = [
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/",
	"http://purl.oclc.org/ooxml/officeDocument/relationships/",
	"http://schemas.microsoft.com/office/2006/relationships/",
	"http://schemas.microsoft.com/office/2007/relationships/",
	"http://schemas.microsoft.com/office/2011/relationships/",
	"http://schemas.microsoft.com/office/2016/09/relationships/",
	"http://schemas.microsoft.com/office/2018/08/relationships/"
]

const STRICT_RELATIONSHIP_BASE = "http://purl.oclc.org/ooxml/officeDocument/relationships/"
const SIGNATURE_ORIGIN = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin"
const SIGNATURE = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"

// The last segment of a known relationship type ("officeDocument", "styles", "header", "commentsIds"...),
// so transitional, Strict and Microsoft extension types compare alike; undefined for unknown types.
export function relationshipKind(type: string): string | undefined {
	for (const base of RELATIONSHIP_BASES) {
		if (type.startsWith(base)) return type.slice(base.length)
	}

	return undefined
}

const MAIN_CONTENT_TYPES = new Map<string, DocxVariant>([
	["application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "docx"],
	["application/vnd.ms-word.document.macroenabled.main+xml", "docm"],
	["application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml", "dotx"],
	["application/vnd.ms-word.template.macroenabledtemplate.main+xml", "dotm"]
])

export class DocxPackage {
	readonly opc: OpcPackage
	readonly contentTypes: ContentTypes
	readonly mainPart: string
	readonly facts: DocxFacts
	readonly settings: DocxSettings
	readonly settingsPart: string | undefined
	private readonly mainRelationships: Relationships
	private readonly documents = new Map<string, Promise<XmlDocument>>()

	constructor(
		opc: OpcPackage,
		contentTypes: ContentTypes,
		mainPart: string,
		mainRelationships: Relationships,
		facts: DocxFacts,
		settings: DocxSettings,
		settingsPart: string | undefined
	) {
		this.opc = opc
		this.contentTypes = contentTypes
		this.mainPart = mainPart
		this.mainRelationships = mainRelationships
		this.facts = facts
		this.settings = settings
		this.settingsPart = settingsPart
	}

	relationships(part: string): Promise<Relationships> {
		return this.opc.relationships(part)
	}

	// Internal parts the main document points to by relationship kind, in relationship order.
	parts(kind: string): string[] {
		const out: string[] = []

		for (const relationship of this.mainRelationships.all) {
			if (
				!relationship.external &&
				relationship.partName !== undefined &&
				relationshipKind(relationship.type) === kind &&
				this.opc.has(relationship.partName)
			) {
				out.push(relationship.partName)
			}
		}

		return out
	}

	part(kind: string): string | undefined {
		return this.parts(kind)[0]
	}

	has(part: string): boolean {
		return this.opc.has(part)
	}

	// A parsed part. The parse is shared by every caller until released; the first caller's `slicing` (host
	// turns while inflating and parsing) applies.
	xml(part: string, slicing?: Slicing): Promise<XmlDocument> {
		let promise = this.documents.get(part)

		if (promise === undefined) {
			promise = this.opc.xml(part, slicing)
			this.documents.set(part, promise)
		}

		return promise
	}

	release(part: string): void {
		this.documents.delete(part)
	}
}

export type DocxOpenResult =
	| { readonly ok: true; readonly package: DocxPackage }
	| { readonly ok: false; readonly reason: DocxOpenFailure; readonly message: string }

const OLE_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]

function failure(reason: DocxOpenFailure, message: string): DocxOpenResult {
	return { ok: false, reason, message }
}

export interface DocxPackageOptions extends ZipOpenOptions {
	// Host turns while the settings part is inflated and parsed.
	readonly slicing?: Slicing
}

// Opens a word-processing package: finds the main document through the package relationships, decides
// the variant from its content type and gathers the gate-0 facts (signatures, protection, Strict). Nothing
// but the content types, relationships and settings is read.
export async function openDocxPackage(source: ByteSource, options: DocxPackageOptions = {}): Promise<DocxOpenResult> {
	if (source.size >= 8) {
		const head = await source.read(0, 8)

		if (OLE_SIGNATURE.every((byte, k) => head[k] === byte)) {
			const kind = await compoundFileKind(source)

			return failure(
				kind,
				kind === "encrypted"
					? "Password-protected document"
					: kind === "legacyBinary"
						? "Legacy binary Office file"
						: "OLE2 compound file"
			)
		}
	}

	let opc: OpcPackage

	try {
		opc = await OpcPackage.open(source, options)
	} catch (error) {
		if (error instanceof ZipError) {
			switch (error.code) {
				case "notZip":
					return failure("notZip", error.message)
				case "tooManyEntries":
				case "tooLarge":
					return failure("tooLarge", error.message)
				default:
					return failure("invalidPackage", error.message)
			}
		}

		if (error instanceof OpcError) return failure("invalidPackage", error.message)

		throw error
	}

	try {
		return await openWordprocessing(opc, options.slicing)
	} catch (error) {
		if (error instanceof OpcError || error instanceof XmlError || error instanceof ZipError) {
			return failure(error instanceof ZipError && error.code === "tooLarge" ? "tooLarge" : "invalidPackage", error.message)
		}

		throw error
	}
}

async function openWordprocessing(opc: OpcPackage, slicing: Slicing | undefined): Promise<DocxOpenResult> {
	const [contentTypes, packageRelationships] = await Promise.all([opc.contentTypes(), opc.relationships("")])
	const main = packageRelationships.all.find(
		relationship => !relationship.external && relationshipKind(relationship.type) === "officeDocument"
	)

	if (main?.partName === undefined || !opc.has(main.partName)) return failure("notWordprocessing", "The package has no main document")

	const mainPart = opc.entry(main.partName)?.name ?? main.partName
	const contentType = contentTypes.typeOf(mainPart)?.toLowerCase()
	let variant = contentType === undefined ? undefined : MAIN_CONTENT_TYPES.get(contentType)

	if (variant === undefined) {
		if (
			contentType !== undefined &&
			(contentType.includes("spreadsheetml") ||
				contentType.includes("presentationml") ||
				contentType.includes("ms-excel") ||
				contentType.includes("ms-powerpoint"))
		) {
			return failure("notWordprocessing", `The main part is ${contentType}`)
		}

		variant = "docx"
	}

	const mainRelationships = await opc.relationships(mainPart)
	const kinds = (kind: string) => mainRelationships.all.filter(relationship => relationshipKind(relationship.type) === kind)
	const settingsRelationship = kinds("settings").find(
		relationship => !relationship.external && relationship.partName !== undefined && opc.has(relationship.partName)
	)
	const settingsPart = settingsRelationship?.partName
	let malformed = false
	let settings = NO_SETTINGS
	let settingsRelationships: Relationships | undefined

	if (settingsPart !== undefined) {
		try {
			settings = readSettings(await opc.xml(settingsPart, slicing))
			malformed = settings.malformed
			settingsRelationships = await opc.relationships(settingsPart)
		} catch (error) {
			if (!(error instanceof XmlError || error instanceof OpcError || error instanceof ZipError)) throw error

			malformed = true
		}
	}

	const macros = variant === "docm" || variant === "dotm" || kinds("vbaProject").length > 0
	const strict = main.type.startsWith(STRICT_RELATIONSHIP_BASE) || settings.strict
	const facts: DocxFacts = {
		variant: macros && variant === "docx" ? "docm" : macros && variant === "dotx" ? "dotm" : variant,
		strict,
		macros,
		template: variant === "dotx" || variant === "dotm",
		signatures: await signatureCount(opc, packageRelationships),
		documentProtection: settings.documentProtection,
		writeProtection: settings.writeProtection,
		glossary: kinds("glossaryDocument").length > 0,
		attachedTemplate:
			settings.attachedTemplate !== undefined ||
			(settingsRelationships?.all.some(relationship => relationshipKind(relationship.type) === "attachedTemplate") ?? false),
		subDocuments: kinds("subDocument").length,
		malformed
	}

	return { ok: true, package: new DocxPackage(opc, contentTypes, mainPart, mainRelationships, facts, settings, settingsPart) }
}

async function signatureCount(opc: OpcPackage, packageRelationships: Relationships): Promise<number> {
	let count = 0

	for (const origin of packageRelationships.byType(SIGNATURE_ORIGIN)) {
		if (origin.external || origin.partName === undefined || !opc.has(origin.partName)) continue

		const relationships = await opc.relationships(origin.partName)

		count += relationships.all.filter(relationship => relationship.type === SIGNATURE && !relationship.external).length
	}

	return count
}

const MAX_DIRECTORY_SECTORS = 4096
const MAX_DIFAT_SECTORS = 1024

// Tells an encrypted OOXML file from a legacy binary one by the stream names in the compound file's
// directory (MS-CFB). Only the header, the FAT sectors the directory chain needs and the directory itself
// are read; anything malformed is just "compoundFile".
async function compoundFileKind(source: ByteSource): Promise<"encrypted" | "legacyBinary" | "compoundFile"> {
	try {
		if (source.size < 512) return "compoundFile"

		const header = new DataView((await source.read(0, 512)).slice().buffer)
		const shift = header.getUint16(0x1e, true)

		if (shift !== 9 && shift !== 12) return "compoundFile"

		const sectorSize = 1 << shift
		const perSector = sectorSize / 4
		const difat: number[] = []

		for (let k = 0; k < 109; k++) difat.push(header.getUint32(0x4c + k * 4, true))

		const sectors = new Map<number, DataView>()
		const sector = async (index: number): Promise<DataView | undefined> => {
			const offset = (index + 1) * sectorSize

			if (index >= 0xfffffffa || offset + sectorSize > source.size) return undefined

			let view = sectors.get(index)

			if (view === undefined) {
				view = new DataView((await source.read(offset, sectorSize)).slice().buffer)
				sectors.set(index, view)
			}

			return view
		}
		// FAT sectors past the header's 109 are listed in a chain of DIFAT sectors: each holds perSector - 1
		// FAT sector numbers and the next DIFAT sector last.
		let difatNext = header.getUint32(0x44, true)
		let difatLeft = Math.min(header.getUint32(0x48, true), MAX_DIFAT_SECTORS)
		const fatSector = async (index: number): Promise<number> => {
			while (index >= difat.length && difatLeft > 0 && difatNext < 0xfffffffa) {
				const chain = await sector(difatNext)

				difatLeft--

				if (chain === undefined) break

				for (let k = 0; k < perSector - 1; k++) difat.push(chain.getUint32(k * 4, true))

				difatNext = chain.getUint32((perSector - 1) * 4, true)
			}

			return difat[index] ?? 0xffffffff
		}
		const names = new Set<string>()
		let current = header.getUint32(0x30, true)

		for (let visited = 0; visited < MAX_DIRECTORY_SECTORS && current < 0xfffffffa; visited++) {
			const directory = await sector(current)

			if (directory === undefined) break

			for (let entry = 0; entry < sectorSize; entry += 128) {
				const length = Math.min(directory.getUint16(entry + 0x40, true), 64)
				const type = directory.getUint8(entry + 0x42)

				if (type === 0 || length < 2) continue

				let name = ""

				for (let k = 0; k < length - 2; k += 2) name += String.fromCharCode(directory.getUint16(entry + k, true))

				names.add(name)
			}

			const fat = await sector(await fatSector(Math.floor(current / perSector)))

			if (fat === undefined) break

			current = fat.getUint32((current % perSector) * 4, true)
		}

		if (names.has("EncryptionInfo") && names.has("EncryptedPackage")) return "encrypted"
		if (names.has("WordDocument") || names.has("Workbook") || names.has("Book") || names.has("PowerPoint Document"))
			return "legacyBinary"

		return "compoundFile"
	} catch {
		return "compoundFile"
	}
}
