import type { RecordStoreKv } from "@/features/videos/lib/recordStore.logic"
import { decodeValue } from "@/lib/storage/decode"

// An in-memory kv with the adapter's semantics: absent and schema-invalid rows read as null.
export function memoryKv(): { kv: RecordStoreKv; rows: Map<string, unknown>; reads: () => number } {
	const rows = new Map<string, unknown>()
	let reads = 0

	return {
		rows,
		reads: () => reads,
		kv: {
			kvGetJson: (key, schema) => {
				reads++

				return Promise.resolve(rows.has(key) ? decodeValue(rows.get(key), schema, "kv", key) : null)
			},
			kvSetJson: (key, value) => {
				rows.set(key, value)

				return Promise.resolve()
			},
			kvKeys: prefix => Promise.resolve([...rows.keys()].filter(key => key.startsWith(prefix))),
			kvEntriesJson: (prefix, schema) =>
				Promise.resolve(
					[...rows.entries()].flatMap(([key, value]) => {
						const parsed = key.startsWith(prefix) ? decodeValue(value, schema, "kv", key) : null

						return parsed === null ? [] : [[key, parsed] as const]
					})
				),
			kvDelete: key => {
				rows.delete(key)

				return Promise.resolve()
			},
			kvDeleteIfAt: entries => {
				const deleted: string[] = []

				for (const [key, at] of entries) {
					const value = rows.get(key)

					if (typeof value === "object" && value !== null && "at" in value && value.at === at) {
						rows.delete(key)
						deleted.push(key)
					}
				}

				return Promise.resolve(deleted)
			}
		}
	}
}
