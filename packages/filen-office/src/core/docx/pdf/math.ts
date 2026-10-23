import { isTokenTag, type MathElement, type MathNode } from "../../math/index"

// An equation's text in reading order (the tokens' text, fractions and scripts written linearly), for its
// alternate description and for drawing it where a PDF cannot carry the math layout.

const MAX_LENGTH = 4096

export function mathText(tree: MathElement): string {
	const out: string[] = []
	let length = 0

	const visit = (node: MathNode, depth: number): void => {
		if (length >= MAX_LENGTH || depth > 64) return

		if (typeof node === "string") {
			out.push(node)
			length += node.length

			return
		}

		if (isTokenTag(node.tag)) {
			for (const child of node.children) visit(child, depth + 1)

			return
		}

		const children = node.children
		const part = (index: number): void => {
			const child = children[index]

			if (child !== undefined) visit(child, depth + 1)
		}

		switch (node.tag) {
			case "mfrac":
				out.push("(")
				part(0)
				out.push(")/(")
				part(1)
				out.push(")")

				return
			case "msup":
				part(0)
				out.push("^")
				part(1)

				return
			case "msub":
				part(0)
				out.push("_")
				part(1)

				return
			case "msubsup":
				part(0)
				out.push("_")
				part(1)
				out.push("^")
				part(2)

				return
			case "msqrt":
				out.push("√(")
				for (const child of children) visit(child, depth + 1)
				out.push(")")

				return
			case "mroot":
				out.push("√(")
				part(0)
				out.push(")")

				return
			default:
				for (const child of children) visit(child, depth + 1)
		}
	}

	visit(tree, 0)

	return out.join("").replace(/\s+/g, " ").trim().slice(0, MAX_LENGTH)
}
