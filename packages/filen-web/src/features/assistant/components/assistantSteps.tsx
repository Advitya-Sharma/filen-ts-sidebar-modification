import { useEffect, useId, useState } from "react"
import { useTranslation } from "react-i18next"
import { CheckIcon, ChevronRightIcon, CircleDashedIcon, CircleSlashIcon, HourglassIcon, Loader2Icon, TriangleAlertIcon } from "lucide-react"
import { cn } from "@filen/shared"
import { useAssistantStore } from "@/features/assistant/store/useAssistantStore"
import { stepsExpanded, type TranscriptEntry, type TranscriptStep } from "@/features/assistant/store/transcript"
import { SLOW_STEP_MS, stepDurationMs, stepFaulted, stepsSummary, visibleProgress } from "@/features/assistant/components/assistant.logic"
import { useDurationText, useStepText } from "@/features/assistant/components/useStepText"

type StepsEntry = Extract<TranscriptEntry, { kind: "steps" }>

const GLYPH_CLASS = "size-3.5 shrink-0"

function StepGlyph({ step }: { step: TranscriptStep }) {
	switch (step.status) {
		case "active":
			return (
				<Loader2Icon
					aria-hidden="true"
					className={cn(GLYPH_CLASS, "motion-safe:animate-spin")}
				/>
			)
		case "queued":
			return (
				<CircleDashedIcon
					aria-hidden="true"
					className={GLYPH_CLASS}
				/>
			)
		case "waiting":
			return (
				<HourglassIcon
					aria-hidden="true"
					className={cn(GLYPH_CLASS, "text-foreground")}
				/>
			)
		case "done":
			return (
				<CheckIcon
					aria-hidden="true"
					className={GLYPH_CLASS}
				/>
			)
		case "failed":
			return stepFaulted(step) ? (
				<TriangleAlertIcon
					aria-hidden="true"
					className={cn(GLYPH_CLASS, "text-destructive")}
				/>
			) : (
				<CircleSlashIcon
					aria-hidden="true"
					className={GLYPH_CLASS}
				/>
			)
		case "stopped":
		case "discarded":
			return (
				<CircleSlashIcon
					aria-hidden="true"
					className={GLYPH_CLASS}
				/>
			)
	}
}

// The only ticking clock in the transcript: mounted for the running step alone, so nothing ticks once the
// turn has moved on or ended.
function LiveElapsed({ since }: { since: number }) {
	const duration = useDurationText()
	const [now, setNow] = useState(() => Date.now())

	useEffect(() => {
		const timer = setInterval(() => {
			setNow(Date.now())
		}, 1000)

		return () => {
			clearInterval(timer)
		}
	}, [])

	const elapsed = now - since

	return elapsed < SLOW_STEP_MS ? null : duration(elapsed, true)
}

function StepRow({ step, animate }: { step: TranscriptStep; animate: boolean }) {
	const stepText = useStepText()
	const duration = useDurationText()
	const { label, detail } = stepText(step)
	const ms = stepDurationMs(step)
	const progress = step.status === "active" ? visibleProgress(step.progress) : null
	const failed = stepFaulted(step)

	return (
		<li
			className={cn(
				"flex min-w-0 items-center gap-2",
				animate && "duration-200 motion-safe:animate-in motion-safe:fade-in-0 motion-safe:slide-in-from-top-1"
			)}
		>
			<StepGlyph step={step} />
			<span className={cn("min-w-0 truncate", step.status === "active" || step.status === "waiting" ? "text-foreground" : "")}>
				{label}
			</span>
			{detail === null ? null : (
				<span className={cn("shrink-0", failed ? "text-destructive" : "")}>
					<span aria-hidden="true">· </span>
					{detail}
				</span>
			)}
			{progress === null ? null : (
				<span
					aria-hidden="true"
					className="h-1 w-10 shrink-0 overflow-hidden rounded-full bg-muted"
				>
					<span
						className="block h-full w-full origin-left rounded-full bg-primary/70 motion-safe:transition-transform motion-safe:duration-200"
						style={{ transform: `scaleX(${String(progress)})` }}
					/>
				</span>
			)}
			<span className="ml-auto shrink-0 pl-2 tabular-nums">
				{step.status === "active" ? (
					<LiveElapsed since={step.startedAt} />
				) : ms !== null && ms >= SLOW_STEP_MS ? (
					duration(ms, false)
				) : null}
			</span>
		</li>
	)
}

// A turn's progress list, right below the user's message. Live and open while the turn runs; once it
// completed, one summary line ("Used 3 tools · 4.2 s") that opens the list again. A stopped or failed
// turn stays open, so the step it ended on shows.
export function AssistantSteps({ entry }: { entry: StepsEntry }) {
	const { t } = useTranslation("assistant")
	const duration = useDurationText()
	const listId = useId()
	const live = entry.outcome === null
	const expanded = live || stepsExpanded(entry)

	if (entry.steps.length === 0) {
		return null
	}

	const summary = stepsSummary(entry)
	const summaryDuration = duration(summary.durationMs, false)

	return (
		<div className="px-5 py-1 text-xs leading-5 text-muted-foreground">
			{live ? null : (
				<button
					type="button"
					aria-expanded={expanded}
					aria-controls={expanded ? listId : undefined}
					className="-mx-1 flex max-w-full items-center gap-1 rounded-md px-1 text-left focus-ring transition-colors outline-none hover:text-foreground"
					onClick={() => {
						useAssistantStore.getState().setStepsExpanded(entry.turnId, !expanded)
					}}
				>
					<ChevronRightIcon
						aria-hidden="true"
						className={cn(
							GLYPH_CLASS,
							"motion-safe:transition-transform motion-safe:duration-150",
							expanded ? "rotate-90" : ""
						)}
					/>
					{summary.failed ? (
						<>
							<TriangleAlertIcon
								aria-hidden="true"
								className={cn(GLYPH_CLASS, "text-destructive")}
							/>
							<span className="sr-only">{t("assistantStepsSummaryFailed")}</span>
						</>
					) : null}
					<span className="min-w-0 truncate">
						{summary.tools > 0
							? t("assistantStepsSummaryTools", { count: summary.tools, duration: summaryDuration })
							: t("assistantStepsSummary", { duration: summaryDuration })}
					</span>
				</button>
			)}
			{expanded ? (
				<ol
					id={listId}
					aria-label={t("assistantStepsLabel")}
					className={cn("flex flex-col gap-0.5", live ? "" : "mt-0.5 ml-1.5 border-l border-border pl-3")}
				>
					{entry.steps.map(step => (
						<StepRow
							key={step.id}
							step={step}
							animate={live}
						/>
					))}
				</ol>
			) : null}
		</div>
	)
}
