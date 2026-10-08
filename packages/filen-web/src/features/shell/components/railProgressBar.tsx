// A rail entry's progress sliver along the icon's bottom edge; the entry itself must be `relative`.
export function RailProgressBar({ percent }: { percent: number }) {
	return (
		<span
			aria-hidden="true"
			className="absolute inset-x-2 bottom-1 h-[3px] overflow-hidden rounded-full bg-foreground/20"
		>
			<span
				className="block h-full rounded-full bg-primary transition-[width] duration-300 ease-out"
				style={{ width: `${String(percent)}%` }}
			/>
		</span>
	)
}

// The live figure a rail entry's tooltip adds after its label.
export function RailTooltipDetail({ children }: { children: string }) {
	return <span className="text-background/60"> · {children}</span>
}
