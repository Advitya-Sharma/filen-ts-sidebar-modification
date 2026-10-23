import { bytesSource } from "../../core/opc/io"
import { openProtectedWorkbook } from "./protection"
import { SHEET_PASSWORD_JOB, type SheetPasswordVerdict } from "./messages"
import type { SheetWorkerJob } from "./worker"

// The password check a protected workbook's PDF export asks first, in the worker that holds the workbook: the file
// as stored is opened again and unlocked with the password, and what that unlocked is let go at once. The page does
// no crypto, and nothing here keeps the password. "default": the file opens with Excel's default password (it never
// asked for one, so its PDF is not protected).
export function sheetPasswordJob(): Readonly<Record<string, SheetWorkerJob>> {
	return {
		[SHEET_PASSWORD_JOB]: async (args, context): Promise<SheetPasswordVerdict> => {
			const { bytes, password } = args as { readonly bytes: Uint8Array; readonly password: string }
			const opened = await openProtectedWorkbook(bytesSource(bytes), context.signal)

			switch (opened.kind) {
				case "unlocked":
					opened.document.release()

					return "default"
				case "needsPassword": {
					const unlocked = await opened.locked.unlock(password, { signal: context.signal })

					if (unlocked.ok) {
						unlocked.document.release()

						return "accepted"
					}

					return unlocked.reason === "wrongPassword" ? "wrong" : unlocked.reason === "aborted" ? "aborted" : "failed"
				}
				case "aborted":
					return "aborted"
				default:
					return "failed"
			}
		}
	}
}
