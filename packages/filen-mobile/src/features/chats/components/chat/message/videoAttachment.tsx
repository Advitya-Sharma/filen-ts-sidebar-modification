import View from "@/components/ui/view"
import { cn } from "@filen/shared"
import Ionicons from "@expo/vector-icons/Ionicons"
import { useVideoPlayer, VideoView } from "expo-video"
import { PressableScale } from "@/components/ui/pressables"
import { type InternalLinkData, linkedVerdictFile, openAttachmentPreview } from "@/features/chats/utils"
import { recordPlayerVerdict } from "@/lib/playerVerdict"

export const VideoAttachment = ({
	url,
	name,
	maxWidth,
	linked,
	fromSelf
}: {
	url: string
	name: string
	maxWidth: number
	linked?: InternalLinkData
	fromSelf: boolean
}) => {
	const player = useVideoPlayer(url, p => {
		p.loop = false
		p.staysActiveInBackground = false

		// The poster frame loads the file anyway: what the load found about a linked file's kind is kept, so
		// an `.mp4` holding audio alone shows as a file attachment from then on.
		const file = linked?.type === "file" ? linkedVerdictFile(linked.file) : null

		if (file) {
			p.addListener("sourceLoad", loaded => {
				recordPlayerVerdict(file, loaded, p.status === "error")
			})
		}
	})

	const style = {
		width: maxWidth,
		height: maxWidth,
		borderRadius: 16
	}

	return (
		<PressableScale
			className={cn(
				"items-center justify-center rounded-2xl overflow-hidden flex-row",
				fromSelf ? "bg-blue-600" : "bg-background-tertiary"
			)}
			style={style}
			onPress={() => {
				openAttachmentPreview({
					linked,
					url,
					name
				})
			}}
		>
			<View className="absolute z-100 bg-transparent w-full h-full items-center justify-center">
				<Ionicons
					name="play-circle-outline"
					size={48}
					color="#ffffff"
				/>
			</View>
			<VideoView
				style={style}
				player={player}
				contentFit="cover"
				nativeControls={false}
				allowsPictureInPicture={false}
				focusable={false}
				fullscreenOptions={{
					enable: false
				}}
			/>
		</PressableScale>
	)
}

export default VideoAttachment
