import type { LayerKey } from "./sizeLayer"

export function layerKeyFor(stable: string | undefined, documentKey: string): LayerKey {
	return stable === undefined ? { kind: "session", id: documentKey } : { kind: "stable", id: stable }
}
