// @vitest-environment happy-dom
import { renderToStaticMarkup } from "react-dom/server"
import { describe, expect, it } from "vitest"
import { markdownUrlTransform } from "../core/markdown/index"
import { MarkdownView, type MarkdownComponents } from "../react/markdown/index"

function html(text: string, components?: MarkdownComponents): string {
	return renderToStaticMarkup(
		<MarkdownView
			text={text}
			components={components}
		/>
	)
}

describe("MarkdownView", () => {
	it("renders CommonMark and GFM without any styling", () => {
		expect(html("# Title\n\n**b** ~~s~~\n\n| a | b |\n|---|---|\n| 1 | 2 |\n\n- [x] done")).toBe(
			"<h1>Title</h1>\n<p><strong>b</strong> <del>s</del></p>\n" +
				"<table><thead><tr><th>a</th><th>b</th></tr></thead><tbody><tr><td>1</td><td>2</td></tr></tbody></table>\n" +
				'<ul class="contains-task-list">\n<li class="task-list-item"><input type="checkbox" disabled="" checked=""/> done</li>\n</ul>'
		)
	})

	it("keeps raw HTML in the source inert", () => {
		const out = html('<script>alert(1)</script>\n\n<img src=x onerror="alert(1)">')

		expect(out).not.toContain("<script")
		expect(out).not.toContain("<img")
	})

	it("drops link and image URLs the shared link policy blocks", () => {
		expect(html("[a](javascript:alert(1)) [b](data:text/html,x) [c](docs/x.md) [d](#top) ![e](javascript:x)")).toBe(
			'<p><a>a</a> <a>b</a> <a>c</a> <a>d</a> <img alt="e"/></p>'
		)
	})

	it("keeps the allowed schemes and autolinks", () => {
		expect(html("[a](https://example.com/A?b=C) [m](mailto:x@example.com) [t](tel:+15551234) https://example.org")).toBe(
			'<p><a href="https://example.com/A?b=C">a</a> <a href="mailto:x@example.com">m</a> <a href="tel:+15551234">t</a> ' +
				'<a href="https://example.org">https://example.org</a></p>'
		)
	})

	it("renders the host's element slots", () => {
		const components: MarkdownComponents = { a: ({ href, children }) => <span data-href={href}>{children}</span> }

		expect(html("[a](https://example.com)", components)).toBe('<p><span data-href="https://example.com">a</span></p>')
	})

	it("takes the host's urlTransform", () => {
		expect(
			renderToStaticMarkup(
				<MarkdownView
					text="[a](https://example.com)"
					urlTransform={() => null}
				/>
			)
		).toBe("<p><a>a</a></p>")
	})
})

describe("markdownUrlTransform", () => {
	it("returns the url unchanged when the policy allows it, else undefined", () => {
		expect(markdownUrlTransform("https://example.com")).toBe("https://example.com")
		expect(markdownUrlTransform("sms:+1")).toBe("sms:+1")
		expect(markdownUrlTransform("javascript:alert(1)")).toBeUndefined()
		expect(markdownUrlTransform("#x")).toBeUndefined()
	})
})
