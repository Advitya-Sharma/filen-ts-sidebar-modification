// Resolves in a task of its own, after whatever else is queued (a message, unlike a timer, is never clamped), so a
// long run of work can let the thread answer in between.
export function nextTask(): Promise<void> {
	return new Promise(resolve => {
		const { port1, port2 } = new MessageChannel()

		port1.onmessage = () => {
			port1.close()
			resolve()
		}
		port2.postMessage(null)
	})
}
