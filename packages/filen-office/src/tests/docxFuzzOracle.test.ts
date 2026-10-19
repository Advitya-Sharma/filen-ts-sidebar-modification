import { describe, expect, it } from "vitest"
import { repeatedIds } from "./fuzz/independentReader"

const encode = (xml: string): Uint8Array => new TextEncoder().encode(xml)
const control = (id: string): string => `<w:sdt><w:sdtPr><w:id w:val="${id}"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>`
const textBox = (inner: string): string =>
	`<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>${inner}</w:drawing></mc:Choice><mc:Fallback><w:pict>${inner}</w:pict></mc:Fallback></mc:AlternateContent>`

describe("the paged fuzz's independent reader", () => {
	it("counts ids a text box's fallback repeats from its choice once, as Word writes them", () => {
		expect(repeatedIds(encode(`<w:body>${textBox(control("7"))}</w:body>`))).toEqual(new Map())
		// Nested in another text box's fallback, still the fallback's.
		expect(repeatedIds(encode(`<w:body>${textBox(control("7") + `<mc:Fallback>${control("7")}</mc:Fallback>`)}</w:body>`))).toEqual(
			new Map()
		)
	})

	it("counts ids repeated outside fallbacks", () => {
		expect(repeatedIds(encode(`<w:body>${control("7")}${textBox(control("7"))}</w:body>`))).toEqual(new Map([["sdt:7", 2]]))
		expect(repeatedIds(encode('<w:body><w:p w14:paraId="0A"/><mc:Fallback/><w:p w14:paraId="0A"/></w:body>'))).toEqual(
			new Map([["paraId:0A", 2]])
		)
	})
})
