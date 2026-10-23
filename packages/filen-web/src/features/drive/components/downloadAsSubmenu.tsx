import { Fragment } from "react"
import { useTranslation } from "react-i18next"
import { type DriveItem } from "@/features/drive/lib/item"
import { ACTION_DEFS } from "@/features/drive/lib/actionDefs"
import { DOWNLOAD_AS_TARGETS, downloadAsTargets } from "@/features/drive/lib/downloadAs.logic"
import { startDownloadAs } from "@/features/drive/lib/downloadAs"
import { type DirectoryTreeMenuFamily } from "@/features/drive/components/directoryTreeSubmenu"
import { usePreviewAccessMode, usePreviewCacheScope } from "@/features/preview/lib/accessMode"

// Download as in the item menus, the bulk context menu and (entries only, inside BulkActionMenuButton) the selection
// bar: the formats every one of `items` exports to, in groups. A pick starts the export (lib/downloadAs.ts); the menu
// closes and its toast takes over. The content mounts only while open (Base UI portals).

export interface DownloadAsEntriesProps {
	family: DirectoryTreeMenuFamily
	items: readonly DriveItem[]
}

export function DownloadAsEntries({ family, items }: DownloadAsEntriesProps) {
	const { t } = useTranslation("drive")
	// A public link's files are read as its viewer reads them.
	const accessMode = usePreviewAccessMode()
	const cacheScope = usePreviewCacheScope()
	const { Item, Separator } = family

	return (
		<>
			{downloadAsTargets(items).map((group, index) => (
				<Fragment key={group.join()}>
					{index > 0 ? <Separator /> : null}
					{group.map(target => {
						const { icon: Icon, labelKey } = DOWNLOAD_AS_TARGETS[target]

						return (
							<Item
								key={target}
								onClick={() => {
									void startDownloadAs(items, target, { accessMode, cacheScope })
								}}
							>
								<Icon aria-hidden="true" />
								{t(labelKey)}
							</Item>
						)
					})}
				</Fragment>
			))}
		</>
	)
}

export function DownloadAsSubmenu({ family, items, disabled }: DownloadAsEntriesProps & { disabled: boolean }) {
	const { t } = useTranslation(["drive", "common"])
	const { Sub, SubTrigger, SubContent } = family
	const def = ACTION_DEFS.downloadAs

	return (
		<Sub>
			<SubTrigger
				disabled={disabled}
				title={disabled ? t("common:offlineActionDisabled") : undefined}
			>
				<def.icon aria-hidden="true" />
				{t(def.labelKey)}
			</SubTrigger>
			<SubContent className="max-w-72">
				<DownloadAsEntries
					family={family}
					items={items}
				/>
			</SubContent>
		</Sub>
	)
}
