import { useTranslation } from "react-i18next"
import { MAX_AUTHOR_NAME_LENGTH, normalizeAuthorName } from "@filen/shared"
import { InputDialog } from "@/components/dialogs/inputDialog"

// The author name comments and tracked changes carry: asked the first time a change needs one, and changed in
// the settings. Starts empty (never filled in from the account). `onTurnOffTracking`: asked because track
// changes needs the name, with turning tracking off offered instead.
export function AuthorNameDialog({
	open,
	pending,
	initialValue,
	onOpenChange,
	onSubmit,
	onTurnOffTracking
}: {
	open: boolean
	pending: boolean
	initialValue: string | null
	onOpenChange: (open: boolean) => void
	onSubmit: (name: string) => void
	onTurnOffTracking?: (() => void) | undefined
}) {
	const { t } = useTranslation("preview")

	return (
		<InputDialog
			open={open}
			pending={pending}
			title={t("previewDocxAuthorNameTitle")}
			body={onTurnOffTracking === undefined ? t("previewDocxAuthorNameBody") : t("previewDocxAuthorNameTrackBody")}
			label={t("previewDocxAuthorNameLabel")}
			initialValue={initialValue ?? ""}
			autoComplete="name"
			maxLength={MAX_AUTHOR_NAME_LENGTH}
			submitLabel={t("previewDocxAuthorNameSave")}
			secondaryAction={
				onTurnOffTracking === undefined ? undefined : { label: t("previewDocxTurnOffTracking"), onClick: onTurnOffTracking }
			}
			validate={value => normalizeAuthorName(value) !== undefined}
			onOpenChange={onOpenChange}
			onSubmit={onSubmit}
		/>
	)
}
