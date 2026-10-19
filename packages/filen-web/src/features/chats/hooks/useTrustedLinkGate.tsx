import { useState, type ReactNode } from "react"
import { useTranslation } from "react-i18next"
import { externalLinkDomain, shouldInterceptLinkClick } from "@/features/chats/lib/linkTrust.logic"
import { trustDomain } from "@/features/chats/lib/trustedDomains"
import { useTrustedDomainsQuery } from "@/features/chats/queries/trustedDomains"
import { ConfirmDialog } from "@/components/dialogs/confirmDialog"

export interface TrustedLinkGate {
	// True when `href` leads to a domain not trusted yet: the confirmation is shown and the caller must not
	// open the link itself. False when it may open now (a trusted domain, a Filen link, no host to confirm).
	intercept: (href: string) => boolean
	// The confirmation, rendered by the caller.
	dialog: ReactNode
}

// The one-time-per-domain confirmation before an external link opens (chat messages, documents): one
// persisted trust set, so a domain confirmed anywhere is trusted everywhere. On confirm the link opens
// first and synchronously, inside the button's user activation (a popup blocker honours window.open only
// there, and activation does not survive an await); the trust write follows.
export function useTrustedLinkGate(): TrustedLinkGate {
	const { t } = useTranslation(["chats", "common"])
	const trustedQuery = useTrustedDomainsQuery()
	const [pending, setPending] = useState<{ href: string; domain: string } | null>(null)

	const intercept = (href: string): boolean => {
		const domain = externalLinkDomain(href)

		if (domain === null || !shouldInterceptLinkClick(domain, trustedQuery.data ?? new Set())) {
			return false
		}

		setPending({ href, domain })

		return true
	}

	const dialog = (
		<ConfirmDialog
			open={pending !== null}
			pending={false}
			title={t("chatLinkTrustTitle")}
			body={t("chatLinkTrustBody", { domain: pending?.domain ?? "" })}
			confirmLabel={t("chatLinkTrustConfirm")}
			cancelLabel={t("common:cancel")}
			onOpenChange={open => {
				if (!open) {
					setPending(null)
				}
			}}
			onConfirm={() => {
				if (pending === null) {
					return
				}

				window.open(pending.href, "_blank", "noopener,noreferrer")
				setPending(null)

				void trustDomain(pending.domain).then(() => trustedQuery.refetch())
			}}
		/>
	)

	return { intercept, dialog }
}
