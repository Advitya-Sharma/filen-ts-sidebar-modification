import { useId, useState } from "react"
import { useTranslation } from "react-i18next"
import { toast } from "sonner"
import { activeLibraryOverrides, newSeriesId, rawLibraryOverride } from "@filen/shared/videos"
import { log } from "@/lib/log"
import { errorLabel } from "@/lib/i18n/errorLabel"
import { useIsOnline } from "@/lib/useIsOnline"
import { editLibraryConfig, useLibraryConfigQuery } from "@/features/videos/queries/library"
import {
	NEW_SERIES,
	canSplit,
	correctionEdit,
	draftErrors,
	draftFor,
	hasCorrection,
	isCorrectionKind,
	mergeEdit,
	overrideFromDraft,
	resetEdit,
	sameOverride,
	splitEdit,
	type ConfigEdit,
	type CorrectionDraft,
	type CorrectionSubject
} from "@/features/videos/lib/corrections.logic"
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from "@/components/ui/dialog"
import { Field, FieldError, FieldGroup, FieldLabel } from "@/components/ui/field"
import { Input } from "@/components/ui/input"
import { Select, SelectContent, SelectGroup, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select"
import { Switch } from "@/components/ui/switch"
import { Button } from "@/components/ui/button"
import { Separator } from "@/components/ui/separator"
import { Spinner } from "@/components/ui/spinner"

// Manual corrections for a title or an episode (research §2.10): title and year, shown as a movie, an episode
// (season and number required) or another video, merging a series into another or splitting it, hiding. Saved in
// the synced library configuration; the classification applies them at once and they always win.

export interface SeriesChoice {
	id: string
	title: string
}

interface Option {
	value: string
	label: string
}

function OptionSelect({
	id,
	options,
	value,
	onChange
}: {
	id: string
	options: readonly Option[]
	value: string
	onChange: (value: string) => void
}) {
	return (
		<Select
			items={options}
			value={value}
			onValueChange={next => {
				if (next !== null) {
					onChange(next)
				}
			}}
		>
			<SelectTrigger
				id={id}
				className="w-full"
			>
				<SelectValue />
			</SelectTrigger>
			<SelectContent>
				<SelectGroup>
					{options.map(option => (
						<SelectItem
							key={option.value}
							value={option.value}
						>
							{option.label}
						</SelectItem>
					))}
				</SelectGroup>
			</SelectContent>
		</Select>
	)
}

// Where the page goes after a save: the title the subject is shown as now (null to stay).
function destinationOf(subject: CorrectionSubject, draft: CorrectionDraft): string | null {
	if (subject.original.kind === "series") {
		return null
	}

	if (draft.kind === "episode") {
		return draft.series === NEW_SERIES ? newSeriesId(subject.key.slice(2)) : draft.series.slice(2)
	}

	// A movie's or other video's id is its key's.
	return subject.key.slice(2)
}

export function CorrectionsDialog({
	subject,
	series,
	onClose,
	onMoved
}: {
	subject: CorrectionSubject
	// Series an item can join or a series can be merged into.
	series: readonly SeriesChoice[]
	onClose: () => void
	// The subject is shown as another title now (`id`): the page follows it.
	onMoved: (id: string) => void
}) {
	const { t } = useTranslation(["videos", "common"])
	const isOnline = useIsOnline()
	const ids = {
		title: useId(),
		year: useId(),
		kind: useId(),
		series: useId(),
		season: useId(),
		episode: useId(),
		episodeEnd: useId(),
		merge: useId(),
		hidden: useId()
	}
	const doc = useLibraryConfigQuery().data?.doc
	const current = doc === undefined ? null : (activeLibraryOverrides(doc)[subject.key] ?? null)
	const raw = doc === undefined ? null : rawLibraryOverride(doc, subject.key)
	const isSeries = subject.original.kind === "series"
	const isEpisode = subject.original.kind === "episode"
	const [draft, setDraft] = useState<CorrectionDraft>(() => draftFor(subject, current))
	const [mergeTarget, setMergeTarget] = useState<string>("")
	const [pending, setPending] = useState(false)
	// Problems show once a save was tried, not while the fields are still being filled in.
	const [attempted, setAttempted] = useState(false)
	const errors = draftErrors(draft)
	const shown = attempted ? errors : []
	const others = series.filter(item => !isSeries || item.id !== subject.seriesId)
	const kindOptions: Option[] = [
		{ value: "movie", label: t("detailKindMovie") },
		{ value: "episode", label: t("detailKindEpisode") },
		{ value: "other", label: t("detailKindOther") }
	]
	const seriesOptions: Option[] = [
		...(isEpisode ? [] : [{ value: NEW_SERIES, label: t("detailNewSeries", { title: subject.original.title }) }]),
		...series.map(item => ({ value: `d:${item.id}`, label: item.title }))
	]
	const mergeOptions: Option[] = others.map(item => ({ value: item.id, label: item.title }))
	const splittable = isSeries && canSplit({ doc, subject })

	function update(patch: Partial<CorrectionDraft>): void {
		setDraft(previous => ({ ...previous, ...patch }))
	}

	async function run(edit: ConfigEdit, moved: string | null): Promise<void> {
		setPending(true)

		try {
			await editLibraryConfig(edit)
			onClose()

			if (moved !== null) {
				onMoved(moved)
			}
		} catch (e) {
			log.warn("videos.corrections", "saving a correction failed", e)
			toast.error(errorLabel(e))
		} finally {
			setPending(false)
		}
	}

	function save(): void {
		if (errors.length > 0) {
			setAttempted(true)

			return
		}

		const next = overrideFromDraft({ subject, draft, raw })

		if (sameOverride(next, raw)) {
			onClose()

			return
		}

		void run(correctionEdit(subject.key, next), destinationOf(subject, draft))
	}

	return (
		<Dialog
			open
			onOpenChange={open => {
				if (!open) {
					onClose()
				}
			}}
		>
			<DialogContent className="sm:max-w-lg">
				<DialogHeader>
					<DialogTitle>{isEpisode ? t("detailEditEpisodeTitle") : t("detailEditTitle")}</DialogTitle>
					<DialogDescription>{t("detailEditDescription")}</DialogDescription>
				</DialogHeader>
				<form
					id={ids.title + "-form"}
					className="flex flex-col gap-5"
					onSubmit={event => {
						event.preventDefault()
						save()
					}}
				>
					<FieldGroup>
						<Field>
							<FieldLabel htmlFor={ids.title}>{t("detailFieldTitle")}</FieldLabel>
							<Input
								id={ids.title}
								value={draft.title}
								aria-invalid={shown.includes("title")}
								maxLength={256}
								onChange={event => {
									update({ title: event.target.value })
								}}
							/>
						</Field>
						{!isSeries ? (
							<Field>
								<FieldLabel htmlFor={ids.kind}>{t("detailFieldKind")}</FieldLabel>
								<OptionSelect
									id={ids.kind}
									options={kindOptions}
									value={draft.kind}
									onChange={value => {
										if (isCorrectionKind(value)) {
											update({ kind: value })
										}
									}}
								/>
							</Field>
						) : null}
						{isSeries || draft.kind === "movie" ? (
							<Field>
								<FieldLabel htmlFor={ids.year}>{t("detailFieldYear")}</FieldLabel>
								<Input
									id={ids.year}
									inputMode="numeric"
									value={draft.year}
									aria-invalid={shown.includes("year")}
									aria-describedby={shown.includes("year") ? `${ids.year}-error` : undefined}
									onChange={event => {
										update({ year: event.target.value })
									}}
								/>
								{shown.includes("year") ? (
									<FieldError id={`${ids.year}-error`}>{t("detailFieldYearInvalid")}</FieldError>
								) : null}
							</Field>
						) : null}
						{!isSeries && draft.kind === "episode" ? (
							<>
								<Field>
									<FieldLabel htmlFor={ids.series}>{t("detailFieldSeries")}</FieldLabel>
									<OptionSelect
										id={ids.series}
										options={seriesOptions}
										value={draft.series}
										onChange={value => {
											update({ series: value })
										}}
									/>
								</Field>
								<div className="grid grid-cols-2 gap-4">
									<Field>
										<FieldLabel htmlFor={ids.season}>{t("detailFieldSeason")}</FieldLabel>
										<Input
											id={ids.season}
											inputMode="numeric"
											aria-required
											value={draft.season}
											aria-invalid={shown.includes("season")}
											onChange={event => {
												update({ season: event.target.value })
											}}
										/>
									</Field>
									<Field>
										<FieldLabel htmlFor={ids.episode}>{t("detailFieldEpisode")}</FieldLabel>
										<Input
											id={ids.episode}
											inputMode="numeric"
											aria-required
											value={draft.episode}
											aria-invalid={shown.includes("episode")}
											onChange={event => {
												update({ episode: event.target.value })
											}}
										/>
									</Field>
								</div>
								<Field>
									<FieldLabel htmlFor={ids.episodeEnd}>{t("detailFieldEpisodeEnd")}</FieldLabel>
									<Input
										id={ids.episodeEnd}
										inputMode="numeric"
										value={draft.episodeEnd}
										aria-invalid={shown.includes("episode")}
										onChange={event => {
											update({ episodeEnd: event.target.value })
										}}
									/>
								</Field>
								{shown.includes("season") || shown.includes("episode") ? (
									<FieldError>{t("detailFieldEpisodeRequired")}</FieldError>
								) : null}
							</>
						) : null}
						{isEpisode ? null : (
							<Field orientation="horizontal">
								<FieldLabel htmlFor={ids.hidden}>{t("detailFieldHidden")}</FieldLabel>
								<Switch
									id={ids.hidden}
									checked={draft.hidden}
									onCheckedChange={hidden => {
										update({ hidden })
									}}
								/>
							</Field>
						)}
					</FieldGroup>
				</form>
				{isSeries && (mergeOptions.length > 0 || splittable) ? (
					<>
						<Separator />
						<div className="flex flex-col gap-3">
							{mergeOptions.length > 0 ? (
								<Field>
									<FieldLabel htmlFor={ids.merge}>{t("detailMergeInto")}</FieldLabel>
									<div className="flex gap-2">
										<div className="min-w-0 flex-1">
											<OptionSelect
												id={ids.merge}
												options={[{ value: "", label: t("detailMergeChoose") }, ...mergeOptions]}
												value={mergeTarget}
												onChange={setMergeTarget}
											/>
										</div>
										<Button
											variant="outline"
											disabled={mergeTarget === "" || pending || !isOnline}
											onClick={() => {
												void run(mergeEdit(subject.seriesId, mergeTarget), mergeTarget)
											}}
										>
											{t("detailMerge")}
										</Button>
									</div>
								</Field>
							) : null}
							{splittable ? (
								<div className="flex items-center justify-between gap-3">
									<p className="text-sm text-muted-foreground">{t("detailSplitDescription")}</p>
									<Button
										variant="outline"
										disabled={pending || !isOnline}
										onClick={() => {
											void run(splitEdit(subject), null)
										}}
									>
										{t("detailSplit")}
									</Button>
								</div>
							) : null}
						</div>
					</>
				) : null}
				<DialogFooter className="sm:justify-between">
					{!hasCorrection(doc, subject) ? (
						<span />
					) : (
						<Button
							variant="ghost"
							disabled={pending || !isOnline}
							onClick={() => {
								void run(resetEdit(subject), isSeries ? null : subject.key.slice(2))
							}}
						>
							{t("detailReset")}
						</Button>
					)}
					<div className="flex flex-col-reverse gap-2 sm:flex-row">
						<Button
							variant="outline"
							onClick={onClose}
						>
							{t("common:cancel")}
						</Button>
						<Button
							type="submit"
							form={ids.title + "-form"}
							disabled={pending || !isOnline}
						>
							{pending ? <Spinner data-icon="inline-start" /> : null}
							{t("detailSave")}
						</Button>
					</div>
				</DialogFooter>
			</DialogContent>
		</Dialog>
	)
}
