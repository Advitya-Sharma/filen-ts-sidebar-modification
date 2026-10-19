import type { MouseEventHandler } from "react"

export const EXTERNAL_LINK_REL = "noopener noreferrer nofollow"

// A new-tab message link showing its own href as the label.
export function ExternalAnchor({
	href,
	className,
	onClick,
	onAuxClick
}: {
	href: string
	className?: string | undefined
	onClick?: MouseEventHandler<HTMLAnchorElement>
	onAuxClick?: MouseEventHandler<HTMLAnchorElement>
}) {
	return (
		<a
			href={href}
			target="_blank"
			rel={EXTERNAL_LINK_REL}
			className={className}
			onClick={onClick}
			onAuxClick={onAuxClick}
		>
			{href}
		</a>
	)
}
