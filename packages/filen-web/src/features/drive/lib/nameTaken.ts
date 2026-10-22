import { driveItemName, nameCollisionKey } from "@filen/shared"
import { queryClient } from "@/queries/client"
import { driveListingQueryOptions, flushListingCreates } from "@/features/drive/queries/drive"

// Whether a file or directory named `name` already sits in the "drive" directory. An upload onto a
// taken name makes a new version of that file, so a create checks here first. Names are unique per
// directory case-insensitively (nameCollisionKey). A listing read this session under the live socket answers from the
// cache, with the creates still queued for it landed first; any other is read again (the listing's own
// freshness policy), and that read is kept.
export async function nameTakenInDirectory(parentUuid: string | null, name: string): Promise<boolean> {
	flushListingCreates()

	const items = await queryClient.query(driveListingQueryOptions("drive", parentUuid))
	const wanted = nameCollisionKey(name)

	return items.some(item => nameCollisionKey(driveItemName(item)) === wanted)
}
