import { useEffect, useState } from "react"
import { useTranslation } from "react-i18next"
import { toast } from "sonner"
import { AuthorNameDialog } from "@/features/office/components/authorNameDialog"
import { loadOfficeSettings, saveAuthorName, saveOfficeAutosave, useOfficeSettings } from "@/features/office/lib/officeSettings"
import { PreferenceToggleRow } from "@/features/settings/components/settingRows"
import { SettingsRow } from "@/features/settings/components/settingsLayout"
import { Button } from "@/components/ui/button"
import { useIsOnline } from "@/lib/useIsOnline"

// Advanced → documents: the author name comments and tracked changes carry, and autosave. Both are synced
// (features/office/lib/officeSettings.ts); the page reads them when it opens.
export function DocumentsRows() {
	const { t } = useTranslation("settings")
	const settings = useOfficeSettings()
	const online = useIsOnline()
	const [editing, setEditing] = useState(false)
	const [pending, setPending] = useState(false)

	useEffect(() => {
		void loadOfficeSettings()
	}, [])

	function submitName(name: string): void {
		setPending(true)
		saveAuthorName(name).then(
			() => {
				setPending(false)
				setEditing(false)
			},
			() => {
				setPending(false)
				toast.error(t("settingsDocumentsSaveFailed"))
			}
		)
	}

	return (
		<>
			<SettingsRow
				label={t("settingsDocumentsAuthorName")}
				description={t("settingsDocumentsAuthorNameDescription")}
			>
				<div className="flex min-w-0 items-center gap-3">
					<span className="truncate text-sm text-muted-foreground">
						{settings.authorName ?? t("settingsDocumentsAuthorNameUnset")}
					</span>
					<Button
						type="button"
						variant="outline"
						disabled={!settings.loaded || !online}
						aria-label={t("settingsDocumentsAuthorName")}
						onClick={() => {
							setEditing(true)
						}}
					>
						{t("settingsDocumentsAuthorNameEdit")}
					</Button>
				</div>
			</SettingsRow>
			<PreferenceToggleRow
				title={t("settingsDocumentsAutosave")}
				description={t("settingsDocumentsAutosaveDescription")}
				checked={settings.autosave}
				disabled={!settings.loaded || !online}
				onCheckedChange={checked => {
					saveOfficeAutosave(checked).catch(() => {
						toast.error(t("settingsDocumentsSaveFailed"))
					})
				}}
			/>
			<AuthorNameDialog
				open={editing}
				pending={pending}
				initialValue={settings.authorName}
				onOpenChange={setEditing}
				onSubmit={submitName}
			/>
		</>
	)
}
