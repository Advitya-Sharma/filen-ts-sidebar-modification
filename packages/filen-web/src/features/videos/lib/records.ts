import { kvDelete, kvEntriesJson, kvGetJson, kvKeys, kvSetJson } from "@/lib/storage/adapter"
import { RecordStore, type RecordStoreKv } from "@/features/videos/lib/recordStore.logic"
import { PROBE_KV_PREFIX, PROBE_MAX_RECORDS, probeRecordSchema } from "@/features/videos/lib/probeService.logic"
import { ART_KV_PREFIX, ART_MAX_RECORDS, artRecordSchema } from "@/features/videos/lib/artwork.logic"
import { FRAME_HINT_KV_PREFIX, FRAME_HINT_MAX_RECORDS, frameHintSchema } from "@/features/videos/lib/frames.logic"

// The Videos feature's per-file kv records, one store per kind.

const kv: RecordStoreKv = { kvGetJson, kvSetJson, kvKeys, kvEntriesJson, kvDelete }

// About a thousand writes between two checks of a store's record count.
const WRITES_PER_SWEEP = 1024

export const probeRecords = new RecordStore({
	prefix: PROBE_KV_PREFIX,
	schema: probeRecordSchema,
	maxRecords: PROBE_MAX_RECORDS,
	memoryRecords: 1024,
	writesPerSweep: WRITES_PER_SWEEP,
	scope: "videos",
	kv
})

export const artRecords = new RecordStore({
	prefix: ART_KV_PREFIX,
	schema: artRecordSchema,
	maxRecords: ART_MAX_RECORDS,
	memoryRecords: 1024,
	writesPerSweep: WRITES_PER_SWEEP,
	scope: "videos",
	kv
})

export const frameHints = new RecordStore({
	prefix: FRAME_HINT_KV_PREFIX,
	schema: frameHintSchema,
	maxRecords: FRAME_HINT_MAX_RECORDS,
	memoryRecords: 256,
	writesPerSweep: WRITES_PER_SWEEP,
	scope: "videos",
	kv
})
