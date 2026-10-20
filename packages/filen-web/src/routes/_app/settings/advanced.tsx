import { createFileRoute } from "@tanstack/react-router"
import { useTranslation } from "react-i18next"
import { SlidersHorizontalIcon } from "lucide-react"
import { TransferConfigRow } from "@/features/settings/components/advanced/transferConfigRow"
import { ArchiveMemoryRow } from "@/features/settings/components/advanced/archiveMemoryRow"
import { UploadsRow } from "@/features/settings/components/advanced/uploadsRow"
import { LogsBlock } from "@/features/settings/components/advanced/logsBlock"
import { AboutRows } from "@/features/settings/components/advanced/aboutRows"
import { ThirdPartyNoticesRow } from "@/features/settings/components/advanced/thirdPartyNoticesRow"
import { DocumentsRows } from "@/features/settings/components/advanced/documentsRows"
import { DOCX_EDITING } from "@/features/preview/lib/docxEditingFlag"
import { SHEET_NEXT } from "@/features/spreadsheet/next/sheetNextFlag"
import { SettingsGroup, SettingsPage } from "@/features/settings/components/settingsLayout"
import { routeHead } from "@/lib/head/routeHead"
import { i18n } from "@/lib/i18n"

export const Route = createFileRoute("/_app/settings/advanced")({
	head: routeHead({ title: () => [i18n.t("settings:settingsSectionAdvanced"), i18n.t("common:settings")] }),
	component: AdvancedPage
})

// What the documents group covers: the editors this build has (both flags are constants, so a build with neither
// drops the group and its strings).
function useDocumentsDescription(): string {
	const { t } = useTranslation(["settings", "preview"])

	if (!SHEET_NEXT) {
		return t("settingsAdvancedDocumentsDescription")
	}

	return t(DOCX_EDITING ? "preview:previewSheetDocumentsDescription" : "preview:previewSheetDocumentsDescriptionSheets")
}

function AdvancedPage() {
	const { t } = useTranslation("settings")
	const documentsDescription = useDocumentsDescription()

	return (
		<SettingsPage
			icon={SlidersHorizontalIcon}
			title={t("settingsSectionAdvanced")}
		>
			<SettingsGroup
				title={t("settingsAdvancedTransferTitle")}
				description={t("settingsAdvancedTransferDescription")}
			>
				<TransferConfigRow />
			</SettingsGroup>
			<SettingsGroup
				title={t("settingsAdvancedArchiveTitle")}
				description={t("settingsAdvancedArchiveDescription")}
			>
				<ArchiveMemoryRow />
			</SettingsGroup>
			{DOCX_EDITING || SHEET_NEXT ? (
				<SettingsGroup
					title={t("settingsAdvancedDocumentsTitle")}
					description={documentsDescription}
				>
					<DocumentsRows />
				</SettingsGroup>
			) : null}
			<SettingsGroup
				title={t("settingsAdvancedUploadsTitle")}
				description={t("settingsAdvancedUploadsDescription")}
			>
				<UploadsRow />
			</SettingsGroup>
			<SettingsGroup
				title={t("settingsLogsTitle")}
				description={t("settingsLogsDescription", { count: 500 })}
			>
				<LogsBlock />
			</SettingsGroup>
			<SettingsGroup title={t("settingsAboutTitle")}>
				<AboutRows />
				<ThirdPartyNoticesRow />
			</SettingsGroup>
		</SettingsPage>
	)
}
