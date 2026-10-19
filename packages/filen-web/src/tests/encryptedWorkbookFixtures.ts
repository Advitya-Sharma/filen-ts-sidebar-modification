import { readFileSync } from "node:fs"
import { join } from "node:path"
import {
	dataSpaceDefinitionStream,
	dataSpaceMapStream,
	dataSpaceVersionStream,
	transformPrimaryStream
} from "@filen/office/core/encryption"
import { agileStreams, encryptedOfficeFile, standardStreams } from "@filen/office/testing/encryption"

// Password-protected workbooks for the spreadsheet tests and e2e, built at test time from the committed
// fixtures/spreadsheet/lo-plain.xlsx by the office package's Node-crypto generator ([MS-OFFCRYPTO] §2.3.4),
// independently of the WebCrypto backend that opens them. Node only. Each is built once per process.
//   agile     Agile (AES-256, SHA-512), 1,000 spins, WORKBOOK_PASSWORD
//   default   the same under Excel's default password ("VelvetSweatshop"): opened without a prompt
//   slow      100,000 spins (the most a silent default-password try spends): long enough to cancel
//   standard  Standard (AES-128, SHA-1): no integrity check, upgraded to Agile on save
//   irm       the package routed through the DRM transform (§2.2.10, §2.2.11): rights-managed, refused

export const WORKBOOK_PASSWORD = "pässwörd 🔒"

export type EncryptedWorkbookName = "agile" | "default" | "slow" | "standard" | "irm"

const PLAIN = join(process.cwd(), "src/tests/fixtures/spreadsheet/lo-plain.xlsx")
const built = new Map<EncryptedWorkbookName, Promise<Uint8Array>>()

export function plainWorkbook(): Uint8Array {
	return new Uint8Array(readFileSync(PLAIN))
}

function build(name: EncryptedWorkbookName): Promise<Uint8Array> {
	const plaintext = plainWorkbook()

	switch (name) {
		case "agile":
			return encryptedOfficeFile(agileStreams(plaintext, WORKBOOK_PASSWORD, { spinCount: 1000 }))
		case "default":
			return encryptedOfficeFile(agileStreams(plaintext, "VelvetSweatshop", { spinCount: 1000 }))
		case "slow":
			return encryptedOfficeFile(agileStreams(plaintext, WORKBOOK_PASSWORD, { spinCount: 100_000 }))
		case "standard":
			return encryptedOfficeFile(standardStreams(plaintext, WORKBOOK_PASSWORD, 128))
		case "irm":
			return encryptedOfficeFile(
				{ encryptedPackage: agileStreams(plaintext, WORKBOOK_PASSWORD, { spinCount: 1000 }).encryptedPackage },
				[],
				[
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
					}
				]
			)
	}
}

// A fresh copy each call (a test may transfer or zero it).
export async function encryptedWorkbook(name: EncryptedWorkbookName): Promise<Uint8Array> {
	let file = built.get(name)

	if (file === undefined) {
		file = build(name)
		built.set(name, file)
	}

	return (await file).slice()
}
