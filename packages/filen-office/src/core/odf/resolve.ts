import type { OdfElement } from "./tree"
import { NO_PROPERTIES, type OdfProperties, type OdfStyle, type OdfStyleSheet, type StyleScope } from "./styles"
import { parseLength, parsePercent } from "./units"

// Longest parent chain followed; longer chains (or cycles) stop there and are reported.
const MAX_CHAIN = 64

// Attributes whose percent value is relative to the same attribute of the parent style (Part 3 §20.190
// fo:font-size, §20.180 style:font-size-asian/-complex, §20.21x fo:margin-*): resolved to a length when the
// parent has one, so a resolved style never holds a percent of something it no longer knows.
const RELATIVE_TO_PARENT: ReadonlySet<string> = new Set([
	"fo:font-size",
	"style:font-size-asian",
	"style:font-size-complex",
	"fo:margin-left",
	"fo:margin-right",
	"fo:margin-top",
	"fo:margin-bottom"
])

export interface ResolvedOdfStyle {
	// The family's default style first, then the ancestors, then the style itself; missing links left out.
	readonly chain: readonly OdfStyle[]
	// The parent chain was cut (cycle or depth): the style resolved from what was read.
	readonly truncated: boolean
	properties(kind: string): OdfProperties
	get(kind: string, attribute: string): string | undefined
}

function formatPoints(points: number): string {
	return `${Math.round(points * 10000) / 10000}pt`
}

function merge(base: OdfProperties, own: OdfProperties): OdfProperties {
	if (own.attributes.size === 0 && own.children.length === 0) return base
	if (base.attributes.size === 0 && base.children.length === 0) return own

	const attributes = new Map(base.attributes)

	for (const [name, value] of own.attributes) {
		if (RELATIVE_TO_PARENT.has(name)) {
			const percent = parsePercent(value)
			const parent = parseLength(base.attributes.get(name))

			if (percent !== undefined && parent !== undefined) {
				attributes.set(name, formatPoints((parent * percent) / 100))

				continue
			}
		}

		attributes.set(name, value)
	}

	// A child element (tab stops, columns, background image) replaces the inherited one whole.
	let children: readonly OdfElement[] = base.children

	if (own.children.length > 0) {
		const replaced = new Set(own.children.map(child => child.name))

		children = [...base.children.filter(child => !replaced.has(child.name)), ...own.children]
	}

	return { attributes, children }
}

class Resolved implements ResolvedOdfStyle {
	readonly chain: readonly OdfStyle[]
	readonly truncated: boolean
	private readonly merged = new Map<string, OdfProperties>()

	constructor(chain: readonly OdfStyle[], truncated: boolean) {
		this.chain = chain
		this.truncated = truncated
	}

	properties(kind: string): OdfProperties {
		let merged = this.merged.get(kind)

		if (merged === undefined) {
			merged = NO_PROPERTIES

			for (const style of this.chain) merged = merge(merged, style.properties.get(kind) ?? NO_PROPERTIES)

			this.merged.set(kind, merged)
		}

		return merged
	}

	get(kind: string, attribute: string): string | undefined {
		return this.properties(kind).attributes.get(attribute)
	}
}

// Formatting properties of a style with inheritance applied (Part 3 §16.2): the family's default style,
// then each parent (always a common style, §19.510) down to the style. Results are cached per style and
// property kind, so a document's thousands of references to a few dozen styles resolve each once.
export class OdfStyleResolver {
	readonly sheet: OdfStyleSheet
	private readonly cache = new Map<string, Resolved>()
	private readonly reported = new Set<string>()

	constructor(sheet: OdfStyleSheet) {
		this.sheet = sheet
	}

	// The style `name` as the `scope` stream's content uses it; undefined (or an unknown name) resolves to the
	// family's default style alone.
	resolve(family: string, name: string | undefined, scope: StyleScope): ResolvedOdfStyle {
		const cacheKey = `${scope}\u0000${family}\u0000${name ?? ""}`
		let resolved = this.cache.get(cacheKey)

		if (resolved !== undefined) return resolved

		const chain: OdfStyle[] = []
		const seen = new Set<OdfStyle>()
		let truncated = false
		let style = name === undefined ? undefined : this.sheet.style(family, name, scope)

		while (style !== undefined) {
			if (seen.has(style) || chain.length >= MAX_CHAIN) {
				truncated = true

				if (!this.reported.has(cacheKey)) {
					this.reported.add(cacheKey)
					this.sheet.issues.push({ kind: "parentCycle", family, name: name ?? "" })
				}

				break
			}

			seen.add(style)
			chain.push(style)
			style = style.parent === undefined || style.parent === "" ? undefined : this.sheet.commonStyle(family, style.parent)
		}

		const defaults = this.sheet.defaults.get(family)

		if (defaults !== undefined) chain.push(defaults)

		chain.reverse()
		resolved = new Resolved(chain, truncated)
		this.cache.set(cacheKey, resolved)

		return resolved
	}

	// The first value of a style attribute (not a property) along the chain, nearest first: style:list-style-name,
	// style:master-page-name, style:data-style-name... An empty value counts (it cancels an inherited one).
	inherited(family: string, name: string | undefined, scope: StyleScope, attribute: string): string | undefined {
		const chain = this.resolve(family, name, scope).chain

		for (let k = chain.length - 1; k >= 0; k--) {
			const value = chain[k]?.attributes.get(attribute)

			if (value !== undefined) return value
		}

		return undefined
	}
}
