import * as Comlink from "comlink"
import type { File as SdkFile, SocketEvent } from "@filen/sdk-rs"
import {
	USER_DICTIONARY_FILE_NAME,
	UserDictionaryVersionError,
	addUserWord,
	emptyUserDictionary,
	mergeUserDictionaries,
	parseUserDictionary,
	removeUserWord,
	sameUserDictionary,
	serializeUserDictionary,
	userWords,
	type UserDictionaryDoc
} from "@filen/spellcheck"
import { sdkApi } from "@/lib/sdk/client"
import { runOp } from "@/lib/actions/outcome"
import { log } from "@/lib/log"
import { registerSocketHandler } from "@/lib/sdk/socket"
import { dotFilenSubdirUuid, findDotFilenSubdir, getDotFilenSubdir, noteDotFilenFile, onDotFilenSubdirForgotten } from "@/lib/dotFilen"

// The user's spelling dictionary, synced end-to-end encrypted as `.filen/Office/user-dictionary.v1.json`
// (the drive encrypts every file). Read once per session when spellchecking first needs it and again only
// after another session saved it (socket events); an addition shows at once and is merged into the file
// as it stands before the upload, so devices and tabs never drop each other's words. Unchanged state is
// never uploaded: every upload is a new file version.

export const OFFICE_DIR_NAME = "Office"

const UPLOAD_DEADLINE_MS = 30_000
// Far above a full list (50k words is about 1.5 MB): a larger file is not downloaded, and not overwritten.
const MAX_FILE_BYTES = 4 * 1024 * 1024
// Socket events arrive in bursts (a save is several): one reload after them.
const RELOAD_DELAY_MS = 500

export interface UserDictionarySdk {
	// `.filen/Office`, created when missing.
	officeDir: () => Promise<string>
	// `.filen/Office` when it exists; creates nothing.
	findOfficeDir: () => Promise<string | null>
	knownOfficeDir: () => string | null
	noteSystemFile: (uuid: string) => void
	listDirectory: (dirUuid: string) => Promise<SdkFile[]>
	download: (file: SdkFile) => Promise<Uint8Array>
	upload: (parent: string, bytes: Uint8Array) => Promise<SdkFile>
}

type DriveSocketEvent = Extract<SocketEvent, { type: "drive" }>

function isDictionaryFile(file: SdkFile): boolean {
	return file.meta.type === "decoded" && file.meta.data.name.toLowerCase() === USER_DICTIONARY_FILE_NAME
}

export class UserDictionaryStore {
	private readonly sdk: UserDictionarySdk
	private readonly now: () => number
	private readonly listeners = new Set<(words: readonly string[]) => void>()
	private doc: UserDictionaryDoc = emptyUserDictionary()
	private wordsValue: readonly string[] = []
	// What the file held when last read or written; the local doc differs from it while a change waits.
	private synced: UserDictionaryDoc = emptyUserDictionary()
	private head: { uuid: string; parent: string } | null = null
	private loaded = false
	private stale = true
	// Bumped by every notice that the file may have changed, so one arriving during a read keeps it stale.
	private staleEpoch = 0
	// The file was written by a newer app: it is read, never overwritten.
	private readOnly = false
	private readonly ownUuids = new Set<string>()
	private reading: Promise<void> | null = null
	private chain: Promise<void> = Promise.resolve()
	private reloadTimer: ReturnType<typeof setTimeout> | undefined
	private generation = 0

	constructor(sdk: UserDictionarySdk, now: () => number = Date.now) {
		this.sdk = sdk
		this.now = now
	}

	words(): readonly string[] {
		return this.wordsValue
	}

	subscribe(listener: (words: readonly string[]) => void): () => void {
		this.listeners.add(listener)

		return () => {
			this.listeners.delete(listener)
		}
	}

	// Reads the file when this session has not, or another session changed it since.
	load(): Promise<void> {
		if (this.loaded && !this.stale) return Promise.resolve()

		this.reading ??= this.read().finally(() => {
			this.reading = null
		})

		return this.reading
	}

	add(word: string): Promise<void> {
		return this.change(doc => addUserWord(doc, word, this.now()))
	}

	remove(word: string): Promise<void> {
		return this.change(doc => removeUserWord(doc, word, this.now()))
	}

	private change(edit: (doc: UserDictionaryDoc) => UserDictionaryDoc): Promise<void> {
		const next = edit(this.doc)

		if (next === this.doc) return Promise.resolve()

		this.setDoc(next)

		return this.enqueue(() => this.flush())
	}

	private enqueue(task: () => Promise<void>): Promise<void> {
		const run = this.chain.then(task)

		// The chain continues past a failure; the caller still hears it.
		this.chain = run.catch(() => undefined)

		return run
	}

	private setDoc(doc: UserDictionaryDoc): void {
		this.doc = doc

		const words = userWords(doc)

		if (words.length === this.wordsValue.length && words.every((word, i) => word === this.wordsValue[i])) return

		this.wordsValue = words

		for (const listener of [...this.listeners]) listener(words)
	}

	private async read(): Promise<void> {
		const generation = this.generation
		const epoch = this.staleEpoch
		const dir = await this.sdk.findOfficeDir()
		let remote = emptyUserDictionary()
		let head: { uuid: string; parent: string; timestamp: number } | null = null
		let readOnly = false

		if (dir !== null) {
			const files = (await this.sdk.listDirectory(dir)).filter(isDictionaryFile)

			// Normally one file; concurrent first saves can leave two, which merge.
			for (const file of files) {
				this.sdk.noteSystemFile(file.uuid)

				if (Number(file.size) > MAX_FILE_BYTES) {
					readOnly = true
					log.warn("spellcheck", "user dictionary file too large to read", { uuid: file.uuid, size: Number(file.size) })

					continue
				}

				try {
					remote = mergeUserDictionaries(remote, parseUserDictionary(new TextDecoder().decode(await this.sdk.download(file))))
				} catch (error) {
					if (error instanceof UserDictionaryVersionError) readOnly = true
					else if (!(error instanceof SyntaxError)) throw error

					log.warn("spellcheck", "user dictionary file not read", { uuid: file.uuid, error })
				}

				const timestamp = Number(file.timestamp)

				if (head === null || timestamp > head.timestamp) head = { uuid: file.uuid, parent: file.parent, timestamp }
			}
		}

		// Signed out (or reset) meanwhile: this answer belongs to nobody.
		if (generation !== this.generation) return

		this.head = head === null ? null : { uuid: head.uuid, parent: head.parent }
		this.readOnly = readOnly
		this.synced = remote
		this.loaded = true
		// Read successfully: current unless another session saved during the read.
		this.stale = epoch !== this.staleEpoch
		this.setDoc(mergeUserDictionaries(this.doc, remote))

		// Words added before this read (or while it ran) still have to reach the file.
		if (!sameUserDictionary(this.doc, remote)) void this.enqueue(() => this.flush()).catch(() => undefined)
	}

	private async flush(): Promise<void> {
		const generation = this.generation

		if (!this.loaded || this.stale) await this.load()
		if (generation !== this.generation || this.readOnly || sameUserDictionary(this.doc, this.synced)) return

		const doc = this.doc
		const parent = this.head?.parent ?? (await this.sdk.officeDir())
		const bytes = new TextEncoder().encode(serializeUserDictionary(doc))
		const file = await this.sdk.upload(parent, bytes)

		if (generation !== this.generation) return

		this.ownUuids.add(file.uuid)
		this.sdk.noteSystemFile(file.uuid)
		this.head = { uuid: file.uuid, parent: file.parent }
		this.synced = doc
	}

	// Another session saved the file (or it went away): read it again, now when someone listens.
	handleDriveEvent(event: DriveSocketEvent): void {
		// Before the first read starts there is nothing to keep fresh.
		if (!this.loaded && this.reading === null) return

		const inner = event.inner

		switch (inner.type) {
			case "fileNew":
			case "fileRestore":
			case "fileMove":
			case "fileArchiveRestored": {
				const ours = inner.file.uuid === this.head?.uuid
				const inDir = inner.file.parent === (this.head?.parent ?? this.sdk.knownOfficeDir())

				if ((inDir && isDictionaryFile(inner.file) && !this.ownUuids.has(inner.file.uuid)) || (ours && !inDir)) this.markStale()

				return
			}
			case "fileTrash":
			case "fileDeletedPermanent":
			case "fileMetadataChanged":
				if (inner.uuid === this.head?.uuid) this.markStale()

				return
			default:
				return
		}
	}

	// Events may have been missed (a socket gap, the directory forgotten).
	markStale(): void {
		this.staleEpoch++
		this.stale = true

		if (!this.loaded || this.listeners.size === 0) return

		clearTimeout(this.reloadTimer)
		this.reloadTimer = setTimeout(() => {
			this.load().catch((error: unknown) => {
				log.warn("spellcheck", "user dictionary reload failed", { error })
			})
		}, RELOAD_DELAY_MS)
	}

	// Sign-out: nothing of this account survives; work in flight answers no one.
	reset(): void {
		this.generation++
		clearTimeout(this.reloadTimer)
		this.loaded = false
		this.stale = true
		this.readOnly = false
		this.head = null
		this.synced = emptyUserDictionary()
		this.ownUuids.clear()
		this.reading = null
		this.chain = Promise.resolve()
		this.setDoc(emptyUserDictionary())
	}
}

export const userDictionary = new UserDictionaryStore({
	officeDir: () => getDotFilenSubdir(OFFICE_DIR_NAME),
	findOfficeDir: () => findDotFilenSubdir(OFFICE_DIR_NAME),
	knownOfficeDir: () => dotFilenSubdirUuid(OFFICE_DIR_NAME),
	noteSystemFile: noteDotFilenFile,
	listDirectory: async dirUuid => (await runOp(sdkApi.listDirectory({ kind: "uuid", uuid: dirUuid }))).files,
	download: file => runOp(sdkApi.downloadFileBytes(file, crypto.randomUUID())),
	upload: (parent, bytes) =>
		runOp(
			sdkApi.uploadFileBytes(parent, Comlink.transfer(bytes, [bytes.buffer]), USER_DICTIONARY_FILE_NAME, "application/json", {
				deadlineMs: UPLOAD_DEADLINE_MS,
				abortToken: crypto.randomUUID()
			})
		)
})

// Called once by the authed shell's socket host.
export function registerUserDictionarySocketHandlers(): () => void {
	const unregister = [
		registerSocketHandler("drive", event => {
			userDictionary.handleDriveEvent(event)
		}),
		registerSocketHandler("driveMalformed", () => {
			userDictionary.markStale()
		}),
		registerSocketHandler("reconnecting", () => {
			userDictionary.markStale()
		}),
		onDotFilenSubdirForgotten(OFFICE_DIR_NAME, () => {
			userDictionary.markStale()
		})
	]

	return () => {
		for (const off of unregister) off()
	}
}
