// Media specs play real audio; nothing asserts on what reaches the speakers. Chromium and Firefox
// mute at the output (BROWSER_USE in playwright.config.ts), which pages cannot observe. WebKit has
// no such switch, so this init script plays every media element at volume 0 while the page keeps
// reading and writing its own volume.
export function silenceMediaElements(): void {
	const volume = Object.getOwnPropertyDescriptor(HTMLMediaElement.prototype, "volume") ?? {}
	const get: unknown = Reflect.get(volume, "get")
	const set: unknown = Reflect.get(volume, "set")
	const play: unknown = Object.getOwnPropertyDescriptor(HTMLMediaElement.prototype, "play")?.value

	if (typeof get !== "function" || typeof set !== "function" || typeof play !== "function") {
		return
	}

	const wanted = new WeakMap<HTMLMediaElement, number>()

	const silence = (element: HTMLMediaElement) => {
		if (Reflect.apply(get, element, []) !== 0) {
			Reflect.apply(set, element, [0])
		}
	}

	Object.defineProperty(HTMLMediaElement.prototype, "volume", {
		configurable: true,
		enumerable: true,
		get(this: HTMLMediaElement) {
			return wanted.get(this) ?? 1
		},
		set(this: HTMLMediaElement, value: number) {
			if (!(value >= 0 && value <= 1)) {
				throw new DOMException(`The volume ${String(value)} is outside [0, 1].`, "IndexSizeError")
			}

			const previous = wanted.get(this) ?? 1

			wanted.set(this, value)
			silence(this)

			if (previous !== value) {
				this.dispatchEvent(new Event("volumechange"))
			}
		}
	})

	HTMLMediaElement.prototype.play = function (this: HTMLMediaElement) {
		silence(this)

		return Reflect.apply(play, this, []) as Promise<void>
	}

	// `autoplay` starts without play(); media events do not bubble, so listen in the capture phase.
	document.addEventListener(
		"loadstart",
		event => {
			if (event.target instanceof HTMLMediaElement) {
				silence(event.target)
			}
		},
		true
	)
}
