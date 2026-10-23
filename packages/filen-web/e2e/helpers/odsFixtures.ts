import { inflateRawSync } from "node:zlib"
import { buildZip } from "./archiveFixtures"

// A small OpenDocument spreadsheet (or template) for the ODS editing legs, and a reader for the package the drive
// then stores.

export const ODS_MIME = "application/vnd.oasis.opendocument.spreadsheet"
export const OTS_MIME = "application/vnd.oasis.opendocument.spreadsheet-template"

const NS =
	'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"'

const text = (value: string): string => `<table:table-cell office:value-type="string"><text:p>${value}</text:p></table:table-cell>`
const number = (value: number): string =>
	`<table:table-cell office:value-type="float" office:value="${String(value)}"><text:p>${String(value)}</text:p></table:table-cell>`

// A row no leg edits: written back byte for byte.
export const ODS_UNTOUCHED_ROW = `<table:table-row>${text("Food")}${number(300)}</table:table-row>`

// "Budget": Item/Cost, Rent 1200, Food 300, a total (SUM over B2:B3, cached 1500); "Notes": one cell.
export function odsFixture(mime: string): Buffer {
	const budget = `<table:table table:name="Budget"><table:table-column table:number-columns-repeated="2"/><table:table-row>${text("Item")}${text("Cost")}</table:table-row><table:table-row>${text("Rent")}${number(1200)}</table:table-row>${ODS_UNTOUCHED_ROW}<table:table-row>${text("Total")}<table:table-cell table:formula="of:=SUM([.B2:.B3])" office:value-type="float" office:value="1500"><text:p>1500</text:p></table:table-cell></table:table-row></table:table>`
	const notes = `<table:table table:name="Notes"><table:table-column/><table:table-row>${text("second sheet")}</table:table-row></table:table>`
	const content = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS} office:version="1.3"><office:automatic-styles/><office:body><office:spreadsheet>${budget}${notes}</office:spreadsheet></office:body></office:document-content>`
	const styles = `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NS} office:version="1.3"><office:styles><style:style style:name="Default" style:family="table-cell"/></office:styles></office:document-styles>`
	const manifest = `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${mime}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/></manifest:manifest>`

	// ODF 1.3 Part 2 §3.3: "mimetype" first and stored.
	return buildZip([
		{ name: "mimetype", data: mime, method: "stored" },
		{ name: "content.xml", data: content },
		{ name: "styles.xml", data: styles },
		{ name: "META-INF/manifest.xml", data: manifest }
	])
}

// An entry of a zip by name: its bytes (inflated where deflated), its place in the file and its method.
export function zipEntry(file: Buffer, name: string): { data: Buffer; index: number; method: number } | undefined {
	let at = file.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]))
	const count = file.readUInt16LE(at + 10)
	const entries: { name: string; offset: number; method: number; size: number }[] = []

	at = file.readUInt32LE(at + 16)

	for (let k = 0; k < count; k++) {
		const nameLength = file.readUInt16LE(at + 28)

		entries.push({
			name: file.subarray(at + 46, at + 46 + nameLength).toString("utf8"),
			offset: file.readUInt32LE(at + 42),
			method: file.readUInt16LE(at + 10),
			size: file.readUInt32LE(at + 20)
		})
		at += 46 + nameLength + file.readUInt16LE(at + 30) + file.readUInt16LE(at + 32)
	}

	entries.sort((a, b) => a.offset - b.offset)

	const index = entries.findIndex(item => item.name === name)
	const found = entries[index]

	if (found === undefined) return undefined

	const start = found.offset + 30 + file.readUInt16LE(found.offset + 26) + file.readUInt16LE(found.offset + 28)
	const raw = file.subarray(start, start + found.size)

	return { data: found.method === 0 ? raw : inflateRawSync(raw), index, method: found.method }
}
