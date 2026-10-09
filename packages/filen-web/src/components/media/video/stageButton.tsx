import type { ComponentProps, ReactNode } from "react"
import { Tooltip as TooltipPrimitive } from "@base-ui/react/tooltip"
import { cn } from "@filen/shared"
import { Button } from "@/components/ui/button"
import { Kbd } from "@/components/ui/kbd"
import { useVideoStageContainer } from "@/components/media/video/stageContext"

type ButtonProps = ComponentProps<typeof Button>

// A tooltip that portals into the video stage (so it shows in fullscreen too).
export function StageTooltip({ label, shortcut, children }: { label: string; shortcut?: string | undefined; children: ReactNode }) {
	const container = useVideoStageContainer()

	return (
		<TooltipPrimitive.Root>
			<TooltipPrimitive.Trigger render={<span className="inline-flex">{children}</span>} />
			<TooltipPrimitive.Portal container={container}>
				<TooltipPrimitive.Positioner
					side="top"
					sideOffset={6}
					className="isolate z-50"
				>
					<TooltipPrimitive.Popup className="dark inline-flex w-fit items-center gap-1.5 rounded-xl bg-neutral-950/90 px-2.5 py-1 text-xs text-foreground ring-1 ring-white/15 data-open:animate-in data-open:fade-in-0 data-closed:animate-out data-closed:fade-out-0">
						{label}
						{shortcut === undefined ? null : <Kbd className="bg-white/10 text-foreground/80">{shortcut}</Kbd>}
					</TooltipPrimitive.Popup>
				</TooltipPrimitive.Positioner>
			</TooltipPrimitive.Portal>
		</TooltipPrimitive.Root>
	)
}

// The player's icon button: ghost, labelled, with its shortcut in the tooltip.
export function StageIconButton(props: {
	label: string
	shortcut?: string | undefined
	onClick?: ButtonProps["onClick"]
	pressed?: boolean | undefined
	disabled?: boolean | undefined
	className?: string | undefined
	size?: "icon-sm" | "icon" | "icon-lg"
	children: ReactNode
}) {
	return (
		<StageTooltip
			label={props.label}
			shortcut={props.shortcut}
		>
			<Button
				variant="ghost"
				size={props.size ?? "icon"}
				aria-label={props.label}
				aria-pressed={props.pressed}
				disabled={props.disabled}
				className={cn(
					"text-white hover:bg-white/15 hover:text-white aria-expanded:bg-white/15 aria-expanded:text-white",
					props.className
				)}
				onClick={props.onClick}
			>
				{props.children}
			</Button>
		</StageTooltip>
	)
}
