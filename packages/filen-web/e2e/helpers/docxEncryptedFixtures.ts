import { cfbBytes } from "@filen/office/core/cfb"
import {
	dataSpaceDefinitionStream,
	dataSpaceMapStream,
	dataSpaceVersionStream,
	transformPrimaryStream
} from "@filen/office/core/encryption"
import { agileStreams, encryptedOfficeFile } from "@filen/office/testing/encryption"
import { buildZip } from "./archiveFixtures"

// Password-protected Word documents for the encrypted-document specs, built here byte for byte (no binary
// checked in): a one-paragraph document sealed as Word 2010+ does ([MS-OFFCRYPTO] §2.3.4.10 Agile, AES-256,
// SHA-512, 100,000 spins) by the office package's independent Node-crypto generator inside the compound file
// Word writes, and a rights-managed (IRM) one, whose DataSpaces route the package through the DRM transform
// (§2.2.10, §2.2.11): refused, never asked a password.

export const ENCRYPTED_DOCX_PASSWORD = "Filen e2e pässwörd"
export const ENCRYPTED_DOCX_TEXT = "Encrypted document fixture"

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
const SPIN_COUNT = 100_000

function plainDocx(text: string): Buffer {
	return buildZip([
		{
			name: "[Content_Types].xml",
			data: `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
		},
		{
			name: "_rels/.rels",
			data: `${XML}<Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`
		},
		{
			name: "word/document.xml",
			data: `${XML}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`
		}
	])
}

export async function encryptedDocxBytes(): Promise<Buffer> {
	const streams = agileStreams(plainDocx(ENCRYPTED_DOCX_TEXT), ENCRYPTED_DOCX_PASSWORD, { spinCount: SPIN_COUNT })

	return Buffer.from(await encryptedOfficeFile(streams))
}

export async function rightsManagedDocxBytes(): Promise<Buffer> {
	return Buffer.from(
		await cfbBytes([
			{ path: ["\u0006DataSpaces", "Version"], bytes: dataSpaceVersionStream() },
			{
				path: ["\u0006DataSpaces", "DataSpaceMap"],
				bytes: dataSpaceMapStream([{ components: ["EncryptedPackage"], dataSpace: "DRMEncryptedDataSpace" }])
			},
			{
				path: ["\u0006DataSpaces", "DataSpaceInfo", "DRMEncryptedDataSpace"],
				bytes: dataSpaceDefinitionStream(["DRMEncryptedTransform"])
			},
			{
				path: ["\u0006DataSpaces", "TransformInfo", "DRMEncryptedTransform", "\u0006Primary"],
				bytes: transformPrimaryStream("{C73DFACD-061F-43B0-8B64-0C620D2A8B50}", "Microsoft.Metadata.DRMTransform", false)
			},
			{ path: ["EncryptedPackage"], bytes: new Uint8Array(64) },
			{ path: ["\u0009DRMContent"], bytes: new Uint8Array(32) }
		])
	)
}
