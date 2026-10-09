import type { DirsAndFiles, DriveEvent, File, NonRootDirTagged, SocketEvent, UuidStr } from "@filen/sdk-rs"

// A small in-memory drive for the Videos index tests: directories and files by uuid, walked like
// listDirRecursive, with builders for the drive events that change it.

let counter = 0

// A canonical (lowercase hex) uuid, unique per call.
export function hexUuid(): UuidStr {
	counter++

	return `00000000-0000-4000-8000-${counter.toString(16).padStart(12, "0")}`
}

export interface FileOptions {
	uuid?: UuidStr
	stableUuid?: UuidStr
	size?: number
	timestamp?: number
	key?: string
	version?: 1 | 2 | 3
}

export function driveFile(parent: string, name: string, options: FileOptions = {}): File {
	const uuid = options.uuid ?? hexUuid()
	const size = BigInt(options.size ?? 2_000_000_000)

	return {
		uuid,
		stableUUID: options.stableUuid ?? uuid,
		parent: parent as UuidStr,
		size,
		favorited: false,
		region: "de-1",
		bucket: "filen-1",
		timestamp: BigInt(options.timestamp ?? 1_700_000_000_000),
		chunks: 2n,
		canMakeThumbnail: false,
		meta: {
			type: "decoded",
			data: {
				name,
				mime: "video/x-matroska",
				modified: 1_690_000_000_000n,
				size,
				key: options.key ?? "a".repeat(64),
				version: options.version ?? 3
			}
		}
	}
}

export function driveDir(parent: string, name: string, uuid: UuidStr = hexUuid()): Extract<NonRootDirTagged, { type: "normal" }> {
	return {
		type: "normal",
		uuid,
		parent: parent as UuidStr,
		color: "default",
		timestamp: 0n,
		favorited: false,
		meta: {
			type: "decoded",
			data: {
				name
			}
		}
	}
}

export class FakeDrive {
	public readonly dirs = new Map<string, Extract<NonRootDirTagged, { type: "normal" }>>()
	public readonly files = new Map<string, File>()

	public addDir(parent: string, name: string): string {
		const dir = driveDir(parent, name)

		this.dirs.set(dir.uuid, dir)

		return dir.uuid
	}

	public addFile(parent: string, name: string, options: FileOptions = {}): File {
		const file = driveFile(parent, name, options)

		this.files.set(file.uuid, file)

		return file
	}

	// Everything below `rootUuid`, like listDirRecursive.
	public walk(rootUuid: string): DirsAndFiles {
		const inside = new Set([rootUuid])
		let grew = true

		while (grew) {
			grew = false

			for (const dir of this.dirs.values()) {
				if (!inside.has(dir.uuid) && inside.has(dir.parent)) {
					inside.add(dir.uuid)
					grew = true
				}
			}
		}

		return {
			dirs: [...this.dirs.values()].filter(dir => dir.uuid !== rootUuid && inside.has(dir.uuid)),
			files: [...this.files.values()].filter(file => inside.has(file.parent))
		}
	}
}

export function driveEvent(id: bigint, inner: DriveEvent): Extract<SocketEvent, { type: "drive" }> {
	return {
		type: "drive",
		inner,
		driveMessageId: id
	}
}
