import { useTranslation } from "react-i18next"
import { Dialog as DialogPrimitive } from "@base-ui/react/dialog"
import { XIcon } from "lucide-react"
import { cn } from "@filen/shared"
import { Button } from "@/components/ui/button"
import { Kbd } from "@/components/ui/kbd"
import { STAGE_POPUP_TONE_CLASS, useVideoStageContainer } from "@/components/media/video/stageContext"
import type { VideosKey } from "@/lib/i18n"

// The player's keyboard map (the ? key), over the picture without blurring it.

const SHORTCUTS: readonly { keys: readonly string[]; label: VideosKey }[] = [
	{ keys: ["Space", "K"], label: "playerKeyPlayPause" },
	{ keys: ["←", "J"], label: "playerKeyBack" },
	{ keys: ["→", "L"], label: "playerKeyForward" },
	{ keys: ["↑", "↓"], label: "playerKeyVolume" },
	{ keys: ["M"], label: "playerKeyMute" },
	{ keys: ["F"], label: "playerKeyFullscreen" },
	{ keys: ["C"], label: "playerKeySubtitles" },
	{ keys: ["G", "H"], label: "playerKeySubtitleDelay" },
	{ keys: ["S"], label: "playerKeySkip" },
	{ keys: ["⇧N", "⇧P"], label: "playerKeyNextPrevious" },
	{ keys: ["0–9"], label: "playerKeyPercent" },
	{ keys: ["Home", "End"], label: "playerKeyStartEnd" },
	{ keys: ["PgUp", "PgDn"], label: "playerKeyChapters" },
	{ keys: [",", "."], label: "playerKeyFrame" },
	{ keys: ["<", ">"], label: "playerKeySpeed" },
	{ keys: ["I"], label: "playerKeyPictureInPicture" },
	{ keys: ["?"], label: "playerKeyShortcuts" }
]

export function ShortcutsDialog({ open, onOpenChange }: { open: boolean; onOpenChange: (open: boolean) => void }) {
	const { t } = useTranslation(["videos", "common"])
	const container = useVideoStageContainer()

	return (
		<DialogPrimitive.Root
			open={open}
			onOpenChange={onOpenChange}
		>
			<DialogPrimitive.Portal container={container}>
				<DialogPrimitive.Popup
					data-stage-popup=""
					// Opened by a key, it has no trigger to go back to: the stage keeps the keyboard.
					finalFocus={() => container}
					className={cn(
						"fixed top-1/2 left-1/2 z-50 flex max-h-[min(36rem,calc(100%-2rem))] w-[min(28rem,calc(100%-2rem))] -translate-1/2 flex-col gap-4 overflow-y-auto rounded-3xl p-5 text-sm outline-none data-open:animate-in data-open:fade-in-0 data-closed:animate-out data-closed:fade-out-0",
						STAGE_POPUP_TONE_CLASS
					)}
					// Keys typed in here belong to the dialog, not to the player behind it.
					onKeyDown={event => {
						if (event.key !== "Escape") {
							event.stopPropagation()
						}
					}}
				>
					<div className="flex items-center justify-between">
						<DialogPrimitive.Title className="text-base font-semibold">{t("playerShortcutsTitle")}</DialogPrimitive.Title>
						<DialogPrimitive.Close
							render={
								<Button
									variant="ghost"
									size="icon-sm"
									aria-label={t("common:close")}
								/>
							}
						>
							<XIcon />
						</DialogPrimitive.Close>
					</div>
					<dl className="grid grid-cols-[auto_1fr] gap-x-4 gap-y-2">
						{SHORTCUTS.map(entry => (
							<div
								key={entry.label}
								className="contents"
							>
								<dt className="flex gap-1">
									{entry.keys.map(key => (
										<Kbd
											key={key}
											className="bg-white/10 text-foreground"
										>
											{key}
										</Kbd>
									))}
								</dt>
								<dd className="text-muted-foreground">{t(entry.label)}</dd>
							</div>
						))}
					</dl>
				</DialogPrimitive.Popup>
			</DialogPrimitive.Portal>
		</DialogPrimitive.Root>
	)
}
