import { useQuery, type UseQueryResult } from "@tanstack/react-query"
import { getTrustedDomains } from "@/lib/trustedLinks/trustedDomains"

// Cached read of the persisted trust set, shared by every gated link in the tree. A write (trustDomain)
// is awaited and followed by this query's refetch.
export function useTrustedDomainsQuery(): UseQueryResult<ReadonlySet<string>> {
	return useQuery({
		queryKey: ["chats", "trustedLinkDomains"] as const,
		queryFn: getTrustedDomains
	})
}
