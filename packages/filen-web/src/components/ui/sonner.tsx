import { Toaster as Sonner, type ToasterProps } from "sonner"
import { useTranslation } from "react-i18next"
import { TOAST_ICONS } from "@/components/ui/toastIcons"
import { useTheme } from "@/providers/themeProvider"
import { useToastClearance, useToastSideClearance } from "@/lib/toastClearance"
import {
	TOAST_EDGE_OFFSET_PX,
	TOAST_MOBILE_EDGE_OFFSET_PX,
	TOAST_WIDTH_PX,
	toastBottomOffset,
	toastRightOffset
} from "@/lib/toastClearance.logic"

// Registry default reads the theme from `next-themes`; this app owns its theme provider, so the
// import is repointed and `next-themes` is not a dependency. `theme` is "dark" | "light" | "system",
// exactly Sonner's own theme union — no cast needed.
const Toaster = ({ ...props }: ToasterProps) => {
	const { theme } = useTheme()
	const { t } = useTranslation("common")
	// Flush with the bottom edge unless a registered surface (selection bar, audio player, composer) occupies the
	// toast corner — then lifted by exactly its measured clearance; a panel docked on the right (the
	// assistant) moves the column left by its width the same way. Top/left stay sonner's defaults.
	const clearance = useToastClearance()
	const sideClearance = useToastSideClearance()

	return (
		<Sonner
			theme={theme}
			className="toaster group"
			icons={TOAST_ICONS}
			style={
				{
					"--normal-bg": "var(--popover)",
					"--normal-text": "var(--foreground)",
					"--normal-border": "var(--border)",
					"--border-radius": "var(--radius)",
					"--width": `${String(TOAST_WIDTH_PX)}px`
				} as React.CSSProperties
			}
			offset={{
				right: toastRightOffset({
					sideClearance,
					edge: TOAST_EDGE_OFFSET_PX
				}),
				bottom: toastBottomOffset({
					clearance,
					edge: TOAST_EDGE_OFFSET_PX
				})
			}}
			mobileOffset={{
				bottom: toastBottomOffset({
					clearance,
					edge: TOAST_MOBILE_EDGE_OFFSET_PX
				})
			}}
			// Every toast gets a tabbable dismiss — a timed-only toast is unreachable by keyboard. The
			// label rides inside toastOptions (sonner declares it there, not on ToasterProps) and is not
			// caller-overridable: a caller-supplied toastOptions replaces this object wholesale.
			closeButton
			toastOptions={{
				classNames: {
					toast: "cn-toast"
				},
				closeButtonAriaLabel: t("toastDismiss")
			}}
			{...props}
		/>
	)
}

export { Toaster }
