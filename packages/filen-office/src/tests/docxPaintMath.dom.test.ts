// @vitest-environment happy-dom
import { afterAll, beforeAll, describe, expect, it } from "vitest"
import { MATH_FRAME_ALL, MathFrame, MathTreeError, type MathElement } from "../core/math/index"
import { paintMath } from "../react/docx/paint/math"

const MATHML_NS = "http://www.w3.org/1998/Math/MathML"

// happy-dom creates MathML elements as plain Elements without CSSOM; browsers give them `style`.
const styles = new WeakMap<Element, Record<string, string>>()

beforeAll(() => {
	Object.defineProperty(Element.prototype, "style", {
		configurable: true,
		get(this: Element): Record<string, string> {
			let style = styles.get(this)

			if (style === undefined) {
				style = {}
				styles.set(this, style)
			}

			return style
		}
	})
})

afterAll(() => {
	Reflect.deleteProperty(Element.prototype, "style")
})

function styleOf(element: Element | null | undefined): Record<string, string> {
	return element === null || element === undefined ? {} : (styles.get(element) ?? {})
}

const FRACTION: MathElement = {
	tag: "math",
	attributes: { display: "block", align: "centerGroup", displaystyle: true },
	children: [
		{
			tag: "mrow",
			children: [
				{
					tag: "mfrac",
					attributes: { linethickness: 0 },
					children: [
						{ tag: "mi", attributes: { mathvariant: "normal" }, children: ["x"] },
						{ tag: "mn", children: ["2"] }
					]
				},
				{ tag: "mo", attributes: { stretchy: false, lspace: 0.2777778, rspace: 0, form: "infix" }, children: ["+"] },
				{
					tag: "mstyle",
					attributes: { scriptlevel: 1, color: "FF0000" },
					children: [{ tag: "mtext", children: ["<script>alert(1)</script>"] }]
				},
				{ tag: "mstyle", attributes: { scriptlevel: -2 }, children: [] },
				{ tag: "mspace", attributes: { width: 1.5 }, children: [] }
			]
		}
	]
}

describe("paintMath", () => {
	it("builds MathML elements in tree order", () => {
		const math = paintMath(FRACTION, document)

		expect(math.namespaceURI).toBe(MATHML_NS)
		expect(math.localName).toBe("math")

		const tags = Array.from(math.querySelectorAll("*")).map(element => element.localName)

		expect(tags).toEqual(["mrow", "mfrac", "mi", "mn", "mo", "mstyle", "mtext", "mstyle", "mspace"])

		for (const element of math.querySelectorAll("*")) expect(element.namespaceURI).toBe(MATHML_NS)
	})

	it("serialises typed attributes", () => {
		const math = paintMath(FRACTION, document)

		expect(math.getAttribute("display")).toBe("block")
		expect(math.getAttribute("displaystyle")).toBe("true")
		expect(math.getAttribute("data-fo-align")).toBe("centerGroup")
		expect(math.hasAttribute("align")).toBe(false)
		expect(math.querySelector("mfrac")?.getAttribute("linethickness")).toBe("0em")
		expect(math.querySelector("mi")?.getAttribute("mathvariant")).toBe("normal")

		const operator = math.querySelector("mo")

		expect(operator?.getAttribute("stretchy")).toBe("false")
		expect(operator?.getAttribute("lspace")).toBe("0.278em")
		expect(operator?.getAttribute("rspace")).toBe("0em")
		expect(operator?.getAttribute("form")).toBe("infix")

		const mstyles = math.querySelectorAll("mstyle")

		expect(mstyles[0]?.getAttribute("scriptlevel")).toBe("+1")
		expect(mstyles[0]?.getAttribute("mathcolor")).toBe("#FF0000")
		expect(mstyles[1]?.getAttribute("scriptlevel")).toBe("-2")
		expect(math.querySelector("mspace")?.getAttribute("width")).toBe("1.5em")
	})

	it("puts token text in as text nodes only", () => {
		const math = paintMath(FRACTION, document)
		const text = math.querySelector("mtext")

		expect(text?.textContent).toBe("<script>alert(1)</script>")
		expect(text?.childNodes.length).toBe(1)
		expect(text?.firstChild?.nodeType).toBe(3)
		expect(math.querySelector("script")).toBeNull()
		expect(math.querySelector("mi")?.textContent).toBe("x")
	})

	it("draws frame borders and strikes as inline styles", () => {
		const math = paintMath(
			{
				tag: "math",
				children: [
					{
						tag: "mrow",
						attributes: { frame: MathFrame.TOP | MathFrame.LEFT | MathFrame.STRIKE_HORIZONTAL | MathFrame.STRIKE_UP },
						children: [{ tag: "mi", children: ["a"] }]
					},
					{ tag: "mrow", attributes: { frame: MATH_FRAME_ALL }, children: [] },
					{ tag: "mrow", attributes: { frame: 0 }, children: [] }
				]
			},
			document
		)
		const rows = math.querySelectorAll("mrow")
		const first = styleOf(rows[0])

		expect(first["borderTop"]).toBe("1px solid currentColor")
		expect(first["borderLeft"]).toBe("1px solid currentColor")
		expect(first["borderBottom"]).toBeUndefined()
		expect(first["backgroundImage"]?.split("linear-gradient").length).toBe(3)
		expect(first["backgroundSize"]).toBe("100% 1px, 100% 100%")
		expect(first["backgroundRepeat"]).toBe("no-repeat")

		const all = styleOf(rows[1])

		expect(all["borderRight"]).toBe("1px solid currentColor")
		expect(all["borderBottom"]).toBe("1px solid currentColor")
		expect(all["backgroundSize"]).toBe("100% 1px, 1px 100%, 100% 100%, 100% 100%")
		expect(styleOf(rows[2])).toEqual({})
	})

	it("refuses trees the converter could not have produced", () => {
		const bad =
			(tree: unknown): (() => Element) =>
			() =>
				paintMath(tree as MathElement, document)

		expect(bad({ tag: "script", children: [] })).toThrow(MathTreeError)
		expect(bad({ tag: "math", children: ["loose text"] })).toThrow(MathTreeError)
		expect(bad({ tag: "math", attributes: { onclick: "alert(1)" }, children: [] })).toThrow(MathTreeError)
		expect(bad({ tag: "math", attributes: { color: '000000" onload="x' }, children: [] })).toThrow(MathTreeError)
		expect(bad({ tag: "mi", children: [{ tag: "mi", children: ["x"] }] })).toThrow(MathTreeError)
		expect(bad({ tag: "mfrac", children: [{ tag: "mi", children: ["x"] }] })).toThrow(MathTreeError)
		expect(bad({ tag: "mspace", attributes: { width: Number.POSITIVE_INFINITY }, children: [] })).toThrow(MathTreeError)
	})

	it("handles trees as deep as the limit without recursion", () => {
		let tree: MathElement = { tag: "mi", children: ["x"] }

		for (let depth = 0; depth < 1000; depth++) tree = { tag: "mrow", children: [tree] }

		const math = paintMath({ tag: "math", children: [tree] }, document)

		expect(math.querySelectorAll("mrow").length).toBe(1000)
		expect(math.querySelector("mi")?.textContent).toBe("x")
	})
})
