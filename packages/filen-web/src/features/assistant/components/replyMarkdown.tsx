import { type Components } from "react-markdown"
import { MarkdownRenderer } from "@/features/preview/components/markdownRenderer"

// A reply relays other people's text (summaries of notes, files, chats), so nothing in it is ever a live
// link or a fetched image, whatever spelling got past the shared reply screen: links render as their text,
// images as their alt text, and no URL reaches the DOM.
const REPLY_OVERRIDES: Components = {
	a: ({ children }) => <span className="underline decoration-dotted">{children}</span>,
	img: ({ alt }) => <span>{alt}</span>
}

function noUrl(): null {
	return null
}

export function AssistantReplyMarkdown({ text, alt, className }: { text: string; alt: string; className?: string | undefined }) {
	return (
		<MarkdownRenderer
			text={text}
			alt={alt}
			className={className}
			overrides={REPLY_OVERRIDES}
			urlTransform={noUrl}
		/>
	)
}
