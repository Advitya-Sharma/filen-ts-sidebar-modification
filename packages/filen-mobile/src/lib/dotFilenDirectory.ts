import auth from "@/lib/auth"
import { AnyNormalDir, type Dir } from "@filen/sdk-rs"
import { toSignalOpts } from "@/lib/signals"
import { useSocketStore } from "@/stores/useSocket.store"
import events from "@/lib/events"

// Resolved subdirectories, trusted only within the socket session that resolved them: while it stays
// connected every directory change reaches the drive socket handler, which clears this.
const memo = new Map<string, { sub: Dir; connectedAt: number }>()
// Bumped on every clear, so a resolve that overlapped one never stores what it listed before it.
let generation = 0

export function clearDotFilenDirectoryMemo(): void {
	generation++
	memo.clear()
}

// Local moves, renames, trashes and deletes land before their socket echo.
events.subscribe("driveItemRemoved", clearDotFilenDirectoryMemo)
events.subscribe("driveItemUpdated", clearDotFilenDirectoryMemo)
events.subscribe("logout", clearDotFilenDirectoryMemo)

// Every feature must resolve the same root-level ".filen" directory, so the name match lives here once. Each level is
// found by name hash (one request, no listing or decryption, case-insensitive like the server's hash); only the
// directory handed back is fetched whole, and ".filen" itself only when a missing subdirectory must be created in it.
export async function ensureDotFilenSubdirectory(name: string, signal?: AbortSignal): Promise<Dir> {
	const socket = useSocketStore.getState()
	const memoized = memo.get(name)

	if (socket.state === "connected" && memoized && memoized.connectedAt === socket.connectedAt) {
		return memoized.sub
	}

	const resolvedGeneration = generation
	const connectedAt = socket.connectedAt
	const opts = toSignalOpts(signal)
	const { authedSdkClient } = await auth.getSdkClients()
	const rootUuid = authedSdkClient.root().uuid
	const dotFilenUuid = await authedSdkClient.findDirUuid(rootUuid, ".filen", opts)
	let sub: Dir

	if (dotFilenUuid === undefined) {
		const dotFilenDir = await authedSdkClient.createDir(
			new AnyNormalDir.Root({
				uuid: rootUuid
			}),
			".filen",
			opts
		)

		// Created just now: it has no children to find.
		sub = await authedSdkClient.createDir(new AnyNormalDir.Dir(dotFilenDir), name, opts)
	} else {
		const subUuid = await authedSdkClient.findDirUuid(dotFilenUuid, name, opts)

		sub =
			subUuid !== undefined
				? await authedSdkClient.getDir(subUuid, opts)
				: await authedSdkClient.createDir(new AnyNormalDir.Dir(await authedSdkClient.getDir(dotFilenUuid, opts)), name, opts)
	}

	const after = useSocketStore.getState()

	if (after.state === "connected" && after.connectedAt === connectedAt && generation === resolvedGeneration) {
		memo.set(name, {
			sub,
			connectedAt
		})
	}

	return sub
}
