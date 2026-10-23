import { describe, expect, it } from "vitest"
import { convertText } from "../core/convert/text/index"
import { bytesSource, ZipArchive } from "../core/opc/index"
import { decodeUtf8, encodeUtf8 } from "../core/xml/index"

// RTF form fields (\*\formfield in a FORMTEXT, FORMCHECKBOX or FORMDROPDOWN field's instruction) carry their
// properties into the field's w:ffData (ECMA-376 Part 1 §17.16.17), which Word and LibreOffice read to show them.

async function body(rtf: string): Promise<string> {
	const converted = await convertText(encodeUtf8(`{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}${rtf}}`), "rtf", "docx")

	if (!converted.ok) throw new Error(converted.refusal.message)

	const archive = await ZipArchive.open(bytesSource(converted.bytes))
	const entry = archive.get("word/document.xml")

	return entry === undefined ? "" : decodeUtf8(await archive.read(entry))
}

describe("RTF form fields", () => {
	it("writes a text form field's w:ffData", async () => {
		const xml = await body(
			"\\pard {\\field{\\*\\fldinst { FORMTEXT }{{\\*\\formfield{\\fftype0\\fftypetxt0\\ffmaxlen170{\\*\\ffname Title}{\\*\\ffdeftext Default}}}}}{\\fldrslt {Gambling Act 2005}}}\\par"
		)

		expect(xml).toMatch(
			/<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Title"\/><w:enabled\/><w:calcOnExit w:val="0"\/><w:textInput><w:default w:val="Default"\/><w:maxLength w:val="170"\/><\/w:textInput><\/w:ffData><\/w:fldChar>/u
		)
		expect(xml).toContain("Gambling Act 2005")
	})

	it("writes a check box's and a drop-down's", async () => {
		const xml = await body(
			"\\pard {\\field{\\*\\fldinst { FORMCHECKBOX }{{\\*\\formfield{\\fftype1\\ffres1\\ffdefres0\\ffhps20{\\*\\ffname Check1}}}}}{\\fldrslt }}" +
				"{\\field{\\*\\fldinst { FORMDROPDOWN }{{\\*\\formfield{\\fftype2\\ffres1\\ffdefres0{\\*\\ffname Drop1}{\\*\\ffl One}{\\*\\ffl Two}}}}}{\\fldrslt }}\\par"
		)

		expect(xml).toContain('<w:checkBox><w:size w:val="20"/><w:default w:val="0"/><w:checked w:val="1"/></w:checkBox>')
		expect(xml).toContain(
			'<w:ddList><w:result w:val="1"/><w:default w:val="0"/><w:listEntry w:val="One"/><w:listEntry w:val="Two"/></w:ddList>'
		)
	})

	it("leaves other fields' begin as it was", async () => {
		expect(await body("\\pard {\\field{\\*\\fldinst { PAGE }}{\\fldrslt 1}}\\par")).toContain('<w:fldChar w:fldCharType="begin"/>')
	})

	const field = (inner: string, kind = "FORMTEXT"): string =>
		`\\pard {\\field{\\*\\fldinst { ${kind} }{{\\*\\formfield{${inner}}}}}{\\fldrslt Value}}\\par`

	it("decodes a form field's text in the document's code page, skipping after \\u as \\uc says", async () => {
		expect(await body(`\\ansicpg1252 ${field("\\fftype0{\\*\\ffname Q}{\\*\\ffdeftext \\'93Hi\\'94}")}`)).toContain(
			'<w:default w:val="\u201cHi\u201d"/>'
		)
		expect(await body(`\\ansicpg1251 ${field("\\fftype0{\\*\\ffname \\'cf\\'f0\\'e8}")}`)).toContain(
			'<w:name w:val="\u041f\u0440\u0438"/>'
		)
		expect(await body(`{\\uc0 ${field("\\fftype0{\\*\\ffname A\\u1071 BC}")}}`)).toContain('<w:name w:val="A\u042fBC"/>')
		expect(await body(field("\\fftype0{\\*\\ffname A\\u1071?B}"))).toContain('<w:name w:val="A\u042fB"/>')
	})

	it("cuts a form field's values to the lengths and entries WordprocessingML holds", async () => {
		const text = await body(
			field(
				`\\fftype0{\\*\\ffname ${"N".repeat(80)}}{\\*\\ffentrymcr ${"M".repeat(50)}}{\\*\\ffhelptext ${"H".repeat(300)}}{\\*\\ffstattext ${"S".repeat(200)}}{\\*\\ffdeftext ${"D".repeat(300)}}{\\*\\ffformat ${"F".repeat(80)}}`
			)
		)

		expect(text).toContain(`<w:name w:val="${"N".repeat(20)}"/>`)
		expect(text).toContain(`<w:entryMacro w:val="${"M".repeat(33)}"/>`)
		expect(text).toContain(`<w:helpText w:type="text" w:val="${"H".repeat(255)}"/>`)
		expect(text).toContain(`<w:statusText w:type="text" w:val="${"S".repeat(140)}"/>`)
		expect(text).toContain(`<w:default w:val="${"D".repeat(255)}"/>`)
		expect(text).toContain(`<w:format w:val="${"F".repeat(64)}"/>`)

		const list = await body(
			field(
				`\\fftype2{\\*\\ffname D}{\\*\\ffl ${"E".repeat(300)}}${Array.from({ length: 30 }, (_, k) => `{\\*\\ffl Item${String(k)}}`).join("")}`,
				"FORMDROPDOWN"
			)
		)

		expect(list.split("<w:listEntry ").length - 1).toBe(25)
		expect(list).toContain(`<w:listEntry w:val="${"E".repeat(255)}"/>`)
		expect(await body(field("\\fftype1\\ffhps-5\\ffsize1{\\*\\ffname C}", "FORMCHECKBOX"))).toContain("<w:checkBox><w:sizeAuto/>")
		expect(
			await body(
				field(
					`\\fftype2\\ffres30\\ffdefres28{\\*\\ffname D}${Array.from({ length: 30 }, (_, k) => `{\\*\\ffl Item${String(k)}}`).join("")}`,
					"FORMDROPDOWN"
				)
			)
		).toContain('<w:ddList><w:result w:val="24"/><w:default w:val="24"/>')
		expect(await body(field("\\fftype0\\ffmaxlen2147483647{\\*\\ffname T}"))).toContain('<w:maxLength w:val="32767"/>')
		// Cut by UTF-16 units, never inside a surrogate pair.
		expect(await body(field(`\\fftype0{\\*\\ffname ${"N".repeat(19)}\\u-10179 ?\\u-8704 ?}`))).toContain(
			`<w:name w:val="${"N".repeat(19)}"/>`
		)
	})
})
