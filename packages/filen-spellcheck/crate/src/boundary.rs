//! Everything the JavaScript boundary needs, in safe Rust: the module owns its I/O buffers and its
//! dictionaries, and callers pass only lengths and small integer ids, which are checked here. The caller
//! writes input into the scratch region (`buffer` gives its address) and reads results from the output
//! region (`output`). The `wasm` module only forwards exported functions to a `Boundary`; this one is plain
//! Rust, so the host tests and Miri cover all of it.

#![forbid(unsafe_code)]

use crate::{SEPARATOR, Speller};

/// The id or length is not one the module knows (a dropped dictionary, a stale id).
pub const ERR_ID: i32 = -1;
/// A length reaches past the scratch region.
pub const ERR_RANGE: i32 = -2;
/// The input is not UTF-8.
pub const ERR_UTF8: i32 = -3;
/// The dictionary does not parse.
pub const ERR_PARSE: i32 = -4;
/// The request exceeds the scratch cap or memory ran out.
pub const ERR_MEMORY: i32 = -5;

/// The scratch region never grows past this: the largest dictionary (.aff + .dic) is under 6 MiB.
pub const MAX_SCRATCH: usize = 64 << 20;
/// After a dictionary load a scratch region larger than this is released, so the space serves the next
/// dictionary's parse instead of idling (wasm memory never shrinks).
const KEEP_SCRATCH: usize = 1 << 20;

#[derive(Default)]
pub struct Boundary {
    // Zero-filled when grown: every byte is initialised whatever the caller writes or skips.
    scratch: Vec<u8>,
    output: Vec<u8>,
    // Slot i holds id i + 1. Ids are never reused, so a stale id finds an empty slot, never another
    // dictionary.
    spellers: Vec<Option<Speller>>,
}

fn length(value: u32) -> usize {
    // u32 always fits usize on wasm32 and 64-bit hosts.
    usize::try_from(value).unwrap_or(usize::MAX)
}

fn signed(value: usize) -> i32 {
    i32::try_from(value).unwrap_or(ERR_MEMORY)
}

// Free functions over single fields, so a caller can hold a dictionary and its input while it writes the
// output region.
fn input(scratch: &[u8], offset: usize, len: u32) -> Result<&str, i32> {
    let end = offset.checked_add(length(len)).ok_or(ERR_RANGE)?;
    let bytes = scratch.get(offset..end).ok_or(ERR_RANGE)?;

    core::str::from_utf8(bytes).map_err(|_| ERR_UTF8)
}

fn speller(spellers: &[Option<Speller>], id: u32) -> Result<&Speller, i32> {
    let index = length(id).checked_sub(1).ok_or(ERR_ID)?;

    spellers.get(index).and_then(Option::as_ref).ok_or(ERR_ID)
}

impl Boundary {
    /// The scratch region with room for `min_len` bytes: its address (to write into), or None when that
    /// is over the cap or memory ran out. The address changes when the region grows or is released.
    pub fn buffer(&mut self, min_len: u32) -> Option<usize> {
        let len = length(min_len);

        if len > MAX_SCRATCH {
            return None;
        }

        if self.scratch.len() < len {
            let target = len
                .max(self.scratch.len().saturating_mul(2))
                .min(MAX_SCRATCH);

            self.scratch
                .try_reserve_exact(target - self.scratch.len())
                .ok()?;
            self.scratch.resize(target, 0);
        }

        Some(self.scratch.as_ptr() as usize)
    }

    /// Bytes the scratch region holds now.
    #[must_use]
    pub fn capacity(&self) -> usize {
        self.scratch.len()
    }

    /// The output region's address; valid until the next call.
    #[must_use]
    pub fn output(&self) -> usize {
        self.output.as_ptr() as usize
    }

    /// Parses a dictionary from scratch bytes `[0, aff_len)` (.aff) and `[aff_len, aff_len + dic_len)`
    /// (.dic). Returns its id (> 0) or an error.
    pub fn create(&mut self, aff_len: u32, dic_len: u32) -> i32 {
        let parsed = match (
            input(&self.scratch, 0, aff_len),
            input(&self.scratch, length(aff_len), dic_len),
        ) {
            (Ok(aff), Ok(dic)) => Speller::new(aff, dic).map_err(|_| ERR_PARSE),
            (Err(error), _) | (_, Err(error)) => Err(error),
        };

        if self.scratch.len() > KEEP_SCRATCH {
            self.scratch = Vec::new();
        }

        match parsed {
            Ok(speller) => {
                if self.spellers.try_reserve(1).is_err() {
                    return ERR_MEMORY;
                }

                self.spellers.push(Some(speller));

                signed(self.spellers.len())
            }
            Err(error) => error,
        }
    }

    /// Drops a dictionary. 0, or `ERR_ID` when the id is not live.
    pub fn drop_speller(&mut self, id: u32) -> i32 {
        match length(id)
            .checked_sub(1)
            .and_then(|index| self.spellers.get_mut(index))
        {
            Some(slot @ Some(_)) => {
                *slot = None;

                0
            }
            _ => ERR_ID,
        }
    }

    /// Checks the `SEPARATOR`-separated words in scratch [0, len): one verdict byte per word (1 correct)
    /// in the output region. Returns the word count or an error.
    pub fn check(&mut self, id: u32, len: u32) -> i32 {
        let (speller, words) = match (speller(&self.spellers, id), input(&self.scratch, 0, len)) {
            (Ok(speller), Ok(words)) => (speller, words),
            (Err(error), _) | (_, Err(error)) => return error,
        };
        let count = words.split(SEPARATOR).count();

        self.output.clear();

        if self.output.try_reserve(count).is_err() {
            return ERR_MEMORY;
        }

        self.output.resize(count, 0);

        let checked = speller.check_words(words, &mut self.output);

        signed(checked)
    }

    /// Suggestions for the word in scratch [0, len), `SEPARATOR`-separated in the output region. Returns
    /// their byte length (0: none) or an error. `ngram` 0 runs only the fast edit-based pass.
    pub fn suggest(&mut self, id: u32, len: u32, ngram: u32) -> i32 {
        let suggestions = match (speller(&self.spellers, id), input(&self.scratch, 0, len)) {
            (Ok(speller), Ok(word)) => speller.suggest(word, ngram != 0),
            (Err(error), _) | (_, Err(error)) => return error,
        };

        self.output = Speller::join(&suggestions);

        signed(self.output.len())
    }

    /// Accepts the word in scratch [0, len) from now on. 1 added, 0 not (whitespace, empty), or an error.
    pub fn add(&mut self, id: u32, len: u32) -> i32 {
        let word = match input(&self.scratch, 0, len) {
            Ok(word) => word.to_owned(),
            Err(error) => return error,
        };
        let Some(Some(speller)) = length(id)
            .checked_sub(1)
            .and_then(|index| self.spellers.get_mut(index))
        else {
            return ERR_ID;
        };

        i32::from(speller.add(&word))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const AFF: &[u8] = b"SET UTF-8\nTRY esianrtolcdugmphbyfvkwz\nSFX S Y 1\nSFX S 0 s .\n";
    const DIC: &[u8] = b"2\nhello\nworld/S\n";

    fn write(boundary: &mut Boundary, bytes: &[u8]) -> u32 {
        let len = u32::try_from(bytes.len()).expect("small input");

        boundary.buffer(len).expect("room");
        boundary.scratch[..bytes.len()].copy_from_slice(bytes);

        len
    }

    fn loaded() -> (Boundary, u32) {
        let mut boundary = Boundary::default();
        let len = write(&mut boundary, &[AFF, DIC].concat());
        let aff = u32::try_from(AFF.len()).expect("small");
        let id = boundary.create(aff, len - aff);

        (boundary, u32::try_from(id).expect("an id"))
    }

    #[test]
    fn checks_suggests_and_adds_through_the_regions() {
        let (mut boundary, id) = loaded();
        let len = write(&mut boundary, b"hello\nhelo\nworlds");

        assert_eq!(boundary.check(id, len), 3);
        assert_eq!(boundary.output, [1, 0, 1]);

        let len = write(&mut boundary, b"helo");

        assert!(boundary.suggest(id, len, 0) > 0);
        assert!(boundary.output.starts_with(b"hello"));

        let len = write(&mut boundary, b"zorp");

        assert_eq!(boundary.add(id, len), 1);
        assert_eq!(boundary.check(id, len), 1);
        assert_eq!(boundary.output, [1]);
    }

    #[test]
    fn rejects_unknown_and_stale_ids() {
        let (mut boundary, id) = loaded();
        let len = write(&mut boundary, b"hello");

        for bad in [0, id + 1, u32::MAX] {
            assert_eq!(boundary.check(bad, len), ERR_ID);
            assert_eq!(boundary.suggest(bad, len, 1), ERR_ID);
            assert_eq!(boundary.add(bad, len), ERR_ID);
            assert_eq!(boundary.drop_speller(bad), ERR_ID);
        }

        assert_eq!(boundary.drop_speller(id), 0);
        assert_eq!(boundary.drop_speller(id), ERR_ID);
        assert_eq!(boundary.check(id, len), ERR_ID);

        // A new dictionary never takes a dropped one's id.
        let (aff, all) = (
            u32::try_from(AFF.len()).expect("small"),
            write(&mut boundary, &[AFF, DIC].concat()),
        );

        assert_eq!(
            boundary.create(aff, all - aff),
            i32::try_from(id + 1).expect("small")
        );
    }

    #[test]
    fn rejects_ranges_past_the_scratch_region_and_bad_text() {
        let (mut boundary, id) = loaded();
        let capacity = u32::try_from(boundary.capacity()).expect("small");

        assert_eq!(boundary.check(id, capacity + 1), ERR_RANGE);
        assert_eq!(boundary.check(id, u32::MAX), ERR_RANGE);
        assert_eq!(boundary.create(u32::MAX, u32::MAX), ERR_RANGE);
        assert_eq!(boundary.create(capacity, 1), ERR_RANGE);

        let len = write(&mut boundary, &[0xff, 0xfe]);

        assert_eq!(boundary.check(id, len), ERR_UTF8);
        assert_eq!(boundary.add(id, len), ERR_UTF8);
        assert_eq!(boundary.check(id, 0), 1);
    }

    #[test]
    fn caps_the_scratch_region_and_reports_parse_errors() {
        let mut boundary = Boundary::default();

        assert!(
            boundary
                .buffer(u32::try_from(MAX_SCRATCH + 1).expect("fits"))
                .is_none()
        );
        assert!(boundary.buffer(16).is_some());
        assert_eq!(boundary.capacity(), 16);

        let len = write(&mut boundary, b"FLAG num\n1\nword/abc\n");

        assert_eq!(boundary.create(9, len - 9), ERR_PARSE);
    }

    // Every export with arbitrary ids and lengths answers a value; nothing panics or reads outside.
    #[test]
    fn survives_arbitrary_arguments() {
        let (mut boundary, id) = loaded();
        let mut state = 0x2545_f491_u32;
        let mut next = || {
            state ^= state << 13;
            state ^= state >> 17;
            state ^= state << 5;
            state
        };
        let rounds = if cfg!(miri) { 200 } else { 4_000 };

        for _ in 0..rounds {
            let pick = next();
            let small = next() % 64;
            let any = if pick & 1 == 0 { small } else { next() };
            let some_id = [0, id, id + 1, next()][usize::try_from(pick >> 1 & 3).expect("small")];

            match pick >> 3 & 7 {
                0 => {
                    let _ = boundary.buffer(next() % 128);
                }
                1 => {
                    boundary.check(some_id, any);
                }
                2 => {
                    boundary.suggest(some_id, any.min(small), pick >> 6 & 1);
                }
                3 => {
                    boundary.add(some_id, any);
                }
                4 => {
                    boundary.create(any, small);
                }
                _ => {
                    let _ = boundary.output();
                }
            }
        }
    }
}
