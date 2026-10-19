import { useState, type AnimationEvent } from "react"
import { useTranslation } from "react-i18next"

// A thin bar along the top of the pages while their layout is not final (after a format on a long document the
// pages show the old layout for a while). Its animations time it, so it needs no timer: it shows 400 ms after
// the layout started (one finishing sooner never shows it), and once shown it stays while the layout runs and
// for 300 ms after, then fades, so a relayout just over 400 ms long does not blink. Screen readers have the
// pages' aria-busy; under reduced motion it is a static tint.
export function DocxLayoutBusy({ busy }: { busy: boolean }) {
	const { t } = useTranslation("preview")
	// Its show animation started (the 400 ms passed): it is on screen until it has faded out.
	const [shown, setShown] = useState(false)
	// The layout started again while it was fading out: it stays as it is, without showing anew.
	const [resumed, setResumed] = useState(false)
	const [wasBusy, setWasBusy] = useState(busy)

	if (busy !== wasBusy) {
		setWasBusy(busy)
		setResumed(busy && shown)
	}

	if (!busy && !shown) {
		return null
	}

	const own = (event: AnimationEvent<HTMLDivElement>): boolean => event.target === event.currentTarget

	return (
		<div
			aria-hidden
			title={t("previewDocxLayoutUpdating")}
			data-docx-layout-busy={busy ? "" : "leaving"}
			className={
				busy
					? resumed
						? "absolute inset-x-0 top-0 z-10 h-0.5 overflow-hidden bg-foreground/10 motion-reduce:bg-primary/40"
						: "absolute inset-x-0 top-0 z-10 h-0.5 animate-in overflow-hidden bg-foreground/10 delay-400 fill-mode-backwards fade-in motion-reduce:bg-primary/40"
					: "absolute inset-x-0 top-0 z-10 h-0.5 animate-out overflow-hidden bg-foreground/10 delay-300 duration-200 fill-mode-forwards fade-out motion-reduce:bg-primary/40"
			}
			onAnimationStart={event => {
				if (busy && own(event)) {
					setShown(true)
				}
			}}
			onAnimationEnd={event => {
				if (!busy && own(event)) {
					setShown(false)
				}
			}}
		>
			<div className="absolute inset-y-0 left-0 w-1/3 animate-indeterminate bg-primary motion-reduce:hidden" />
		</div>
	)
}
