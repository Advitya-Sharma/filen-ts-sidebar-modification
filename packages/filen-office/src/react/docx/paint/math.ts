import { MathFrame, assertMathTree, isTokenTag, type MathAttributes, type MathElement } from "../../../core/math/index"

// Builds MathML from a validated math tree: tags and attribute values come from the tree's allowlists and
// are serialised here; document text reaches the DOM only as token textContent.

const MATHML_NS = "http://www.w3.org/1998/Math/MathML"

const BOOLEAN_ATTRIBUTES = [
	"displaystyle",
	"stretchy",
	"symmetric",
	"largeop",
	"movablelimits",
	"fence",
	"separator",
	"accent",
	"accentunder"
] as const

const LENGTH_ATTRIBUTES = ["lspace", "rspace", "minsize", "maxsize", "linethickness", "width", "height", "depth"] as const

const ENUM_ATTRIBUTES = ["display", "mathvariant", "form", "columnalign"] as const

const LINE = "linear-gradient(currentColor, currentColor)"

function diagonal(direction: string): string {
	return `linear-gradient(${direction}, transparent calc(50% - 0.5px), currentColor calc(50% - 0.5px), currentColor calc(50% + 0.5px), transparent calc(50% + 0.5px))`
}

const STRIKE_UP = diagonal("to top right")
const STRIKE_DOWN = diagonal("to bottom right")
const BORDER = "1px solid currentColor"

function length(value: number): string {
	return `${Math.round(value * 1000) / 1000}em`
}

// m:borderBox as CSS borders and strike lines drawn as background layers.
function applyFrame(element: MathMLElement, frame: number): void {
	// MathML elements without CSSOM (engines predating MathML Core) render no math anyway.
	const style: CSSStyleDeclaration | undefined = element.style

	if (frame === 0 || style === undefined) return

	if ((frame & MathFrame.TOP) !== 0) style.borderTop = BORDER
	if ((frame & MathFrame.BOTTOM) !== 0) style.borderBottom = BORDER
	if ((frame & MathFrame.LEFT) !== 0) style.borderLeft = BORDER
	if ((frame & MathFrame.RIGHT) !== 0) style.borderRight = BORDER

	const images: string[] = []
	const sizes: string[] = []

	if ((frame & MathFrame.STRIKE_HORIZONTAL) !== 0) {
		images.push(LINE)
		sizes.push("100% 1px")
	}

	if ((frame & MathFrame.STRIKE_VERTICAL) !== 0) {
		images.push(LINE)
		sizes.push("1px 100%")
	}

	if ((frame & MathFrame.STRIKE_UP) !== 0) {
		images.push(STRIKE_UP)
		sizes.push("100% 100%")
	}

	if ((frame & MathFrame.STRIKE_DOWN) !== 0) {
		images.push(STRIKE_DOWN)
		sizes.push("100% 100%")
	}

	if (images.length > 0) {
		style.backgroundImage = images.join(", ")
		style.backgroundSize = sizes.join(", ")
		style.backgroundPosition = "center"
		style.backgroundRepeat = "no-repeat"
	}

	// Room between the lines and the content.
	style.padding = "0.15em"
}

function applyAttributes(element: MathMLElement, attributes: MathAttributes): void {
	for (const name of BOOLEAN_ATTRIBUTES) {
		const value = attributes[name]

		if (value !== undefined) element.setAttribute(name, value ? "true" : "false")
	}

	for (const name of LENGTH_ATTRIBUTES) {
		const value = attributes[name]

		if (value !== undefined) element.setAttribute(name, length(value))
	}

	for (const name of ENUM_ATTRIBUTES) {
		const value = attributes[name]

		if (value !== undefined) element.setAttribute(name, value)
	}

	const scriptlevel = attributes.scriptlevel

	if (scriptlevel !== undefined) element.setAttribute("scriptlevel", scriptlevel >= 0 ? `+${scriptlevel}` : String(scriptlevel))

	// Not MathML: the caller positions a display equation's block.
	if (attributes.align !== undefined) element.setAttribute("data-fo-align", attributes.align)
	if (attributes.color !== undefined) element.setAttribute("mathcolor", `#${attributes.color}`)
	if (attributes.frame !== undefined) applyFrame(element, attributes.frame)
}

// Throws a MathTreeError for a tree the converter could not have produced, so the caller can show a
// placeholder instead.
export function paintMath(tree: MathElement, document: Document): Element {
	assertMathTree(tree)

	const root = document.createElementNS(MATHML_NS, tree.tag)
	// Iterative: an element and the DOM node its children go into.
	const pending: [MathElement, MathMLElement][] = [[tree, root]]

	for (let next = pending.pop(); next !== undefined; next = pending.pop()) {
		const [node, element] = next

		if (node.attributes !== undefined) applyAttributes(element, node.attributes)

		if (isTokenTag(node.tag)) {
			const text = node.children[0]

			element.textContent = typeof text === "string" ? text : ""
			continue
		}

		for (const child of node.children) {
			if (typeof child === "string") continue

			const childElement = document.createElementNS(MATHML_NS, child.tag)

			element.appendChild(childElement)
			pending.push([child, childElement])
		}
	}

	return root
}
