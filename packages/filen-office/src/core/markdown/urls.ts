import { classifyLink } from "@filen/shared/linkPolicy"

// The rendered view's urlTransform: every link and image URL in the source passes the shared link policy
// or is dropped (undefined omits the attribute). Only an allowed absolute URL survives: a relative or
// fragment-only one would resolve against the host app's own location, and nothing gives headings an id
// for a fragment to target anyway.
export function markdownUrlTransform(url: string): string | undefined {
	return classifyLink(url).kind === "external" ? url : undefined
}
