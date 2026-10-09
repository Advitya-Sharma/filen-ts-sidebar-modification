import { withLibraryRoot, withoutLibraryRoot, type LibraryConfigDoc, type LibraryRootValidation, type Stamp } from "@filen/shared/videos"
import { ROOT_ENTRIES_LARGE } from "@/features/videos/lib/library.logic"

// The "Add a directory" flow: which chooser rows can be picked, and the steps from a pick to the saved root
// (overlap → replace offer, a very large directory → memory warning, then the save).

export interface ChooserRoot {
	uuid: string
	name: string
}

// Filen's own directory at the drive root never shows in the chooser.
export function isHiddenChooserRow({ name, atDriveRoot, systemName }: { name: string; atDriveRoot: boolean; systemName: string }): boolean {
	return atDriveRoot && name.toLowerCase() === systemName
}

// A row that already is a library directory can't be picked again.
export function isRootRow(uuid: string, roots: readonly ChooserRoot[]): boolean {
	return roots.some(root => root.uuid === uuid)
}

// What the chooser says about the open directory before it is added: it is a library directory already, or
// sits inside one (adding it is then offered as a replacement).
export type ChooserPickHint = { kind: "none" } | { kind: "alreadyRoot" } | { kind: "insideRoot"; root: ChooserRoot }

export function chooserPickHint({ pathStack, roots }: { pathStack: readonly string[]; roots: readonly ChooserRoot[] }): ChooserPickHint {
	const current = pathStack.at(-1)

	if (current !== undefined && isRootRow(current, roots)) {
		return {
			kind: "alreadyRoot"
		}
	}

	for (let index = pathStack.length - 2; index >= 0; index--) {
		const root = roots.find(candidate => candidate.uuid === pathStack[index])

		if (root !== undefined) {
			return {
				kind: "insideRoot",
				root
			}
		}
	}

	return {
		kind: "none"
	}
}

export type AddDirectoryStep =
	| { kind: "browse" }
	// The pick is being validated.
	| { kind: "checking" }
	// Validated (replacing `replace`); its size is being read.
	| { kind: "sizing"; uuid: string; replace: string[] }
	// Overlaps these library directories: adding replaces them, once confirmed.
	| { kind: "replace"; uuid: string; replace: string[] }
	| { kind: "large"; uuid: string; replace: string[]; entries: number }
	| { kind: "blocked"; reason: "alreadyRoot" | "systemDirectory" | "notFound" | "notOwnDrive" }
	| { kind: "saving" }

function union(a: readonly string[], b: readonly string[]): string[] {
	return [...new Set([...a, ...b])]
}

// After the pick was validated. `replacing` names directories the pick replaces regardless (the chooser was
// opened to swap an empty one); a validation that finds only those overlapping still counts as clear.
export function stepAfterValidation({
	uuid,
	validation,
	replacing
}: {
	uuid: string
	validation: LibraryRootValidation | { ok: false; reason: "notFound" }
	replacing: readonly string[]
}): AddDirectoryStep {
	const sizing: AddDirectoryStep = {
		kind: "sizing",
		uuid,
		replace: [...replacing]
	}

	if (validation.ok) {
		return sizing
	}

	if (validation.reason !== "nested") {
		return validation.reason === "alreadyRoot" && replacing.includes(uuid)
			? sizing
			: {
					kind: "blocked",
					reason: validation.reason
				}
	}

	const others = validation.replace.filter(root => !replacing.includes(root))

	return others.length === 0
		? sizing
		: {
				kind: "replace",
				uuid,
				replace: union(replacing, validation.replace)
			}
}

// After the pick's size is known (files and directories below it, or null when it couldn't be read).
export function stepAfterSize({
	uuid,
	replace,
	entries
}: {
	uuid: string
	replace: readonly string[]
	entries: number | null
}): AddDirectoryStep {
	return entries !== null && entries >= ROOT_ENTRIES_LARGE
		? {
				kind: "large",
				uuid,
				replace: [...replace],
				entries
			}
		: {
				kind: "saving"
			}
}

// The configuration edit adding `uuid` as a library directory in place of `replace`.
export function addRootEdit({
	uuid,
	replace,
	now
}: {
	uuid: string
	replace: readonly string[]
	now: number
}): (doc: LibraryConfigDoc, stamp: Stamp) => LibraryConfigDoc {
	return (doc, stamp) => {
		let next = doc

		for (const dirUuid of replace) {
			if (dirUuid !== uuid) {
				next = withoutLibraryRoot({
					doc: next,
					dirUuid,
					stamp,
					at: now
				})
			}
		}

		return withLibraryRoot({
			doc: next,
			dirUuid: uuid,
			stamp,
			now
		})
	}
}
