import { verdictChangesKind, type ContentVerdict } from "@filen/shared"
import type { AnyFile } from "@filen/sdk-rs"
import type { VideoProbe } from "@filen/shared/videos"
import { anyFileTypeFacts } from "@/features/drive/lib/item"
import { previewCategoryForExtension } from "@/features/drive/lib/preview.logic"
import { isTransportStreamName } from "@/features/preview/lib/previewPlayback.logic"
import { recordFileKindVerdict } from "@/features/drive/lib/fileKindVerdicts"
import { probeVerdict } from "@/features/drive/lib/fileKindVerdicts.logic"

// The writers of the file-kind verdicts (fileKindVerdicts.ts): a verdict is kept only where it changes what
// the file resolves to, or settles a `.ts` that would otherwise be read again, so the store holds no
// confirmations of what the name already said.

export interface VerdictFile {
	uuid: string
	name: string
	mime: string | null | undefined
	size: number | bigint
}

// `memoryOnly`: see recordFileKindVerdict.
export function recordContentVerdict(file: VerdictFile, verdict: ContentVerdict, options: { memoryOnly?: boolean } = {}): void {
	if (verdictChangesKind(file, verdict, previewCategoryForExtension)) {
		recordFileKindVerdict(file.uuid, verdict, options)
	}
}

// What a container probe that ran anyway says about the file's kind.
export function recordProbeVerdict(
	file: VerdictFile,
	probe: Pick<VideoProbe, "container" | "video" | "audio" | "partial"> | null,
	options: { memoryOnly?: boolean } = {}
): void {
	const verdict = probeVerdict(probe, isTransportStreamName(file.name))

	if (verdict !== null) {
		recordContentVerdict(file, verdict, options)
	}
}

const adopted = new WeakSet<object>()

// Every record a probe service writes leaves its verdict too.
export function adoptProbeRecords(service: {
	onRecord(
		listener: (
			uuid: string,
			record: { probe: Pick<VideoProbe, "container" | "video" | "audio" | "partial"> | null },
			file: AnyFile
		) => void
	): () => void
}): void {
	if (adopted.has(service)) {
		return
	}

	adopted.add(service)
	service.onRecord((uuid, record, file) => {
		const facts = anyFileTypeFacts(file)

		if (facts !== null) {
			recordProbeVerdict({ uuid, ...facts }, record.probe)
		}
	})
}
