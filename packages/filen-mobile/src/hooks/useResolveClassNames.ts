import { useResolveClassNames as useUniwindResolveClassNames } from "uniwind"
import { type NativeColorValue } from "react-native"

// What uniwind actually puts in a resolved color slot: a CSS-compiled color string (or a platform color).
// RN 0.88's `ColorValue` also admits `null | number`, which `@expo/vector-icons` (and other
// `string | OpaqueColorValue` consumers) reject.
export type ResolvedColor = string | NativeColorValue

type UniwindStyle = ReturnType<typeof useUniwindResolveClassNames>

type NarrowColor<T> = NativeColorValue extends T ? Exclude<T, null | number> : T

export type ResolvedStyle = {
	[K in keyof UniwindStyle]: NarrowColor<UniwindStyle[K]>
}

// Type-only narrowing of uniwind's hook — the runtime function is uniwind's own.
export const useResolveClassNames = useUniwindResolveClassNames as (className: string) => ResolvedStyle
