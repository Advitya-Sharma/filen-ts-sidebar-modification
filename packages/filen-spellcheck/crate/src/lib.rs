//! Hunspell-compatible spellchecking for the Filen apps: spellbook behind a small owned API. Native
//! bindings wrap [`Speller`]; the web build exposes it as a C ABI (`wasm` module) that JavaScript calls
//! without generated glue.

#![deny(unsafe_code)]
#![deny(unsafe_op_in_unsafe_fn)]
#![warn(clippy::pedantic)]

use spellbook::Dictionary;

pub mod boundary;
#[cfg(target_arch = "wasm32")]
mod wasm;

/// Why a dictionary could not be loaded.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SpellerError {
    /// The `.aff` or `.dic` text could not be parsed; the message is the parser's.
    Parse(String),
}

impl core::fmt::Display for SpellerError {
    fn fmt(&self, f: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        match self {
            SpellerError::Parse(message) => write!(f, "dictionary parse error: {message}"),
        }
    }
}

impl std::error::Error for SpellerError {}

/// One loaded dictionary.
pub struct Speller {
    dictionary: Dictionary,
}

/// Separates words in batched input and suggestions in output; no word contains it.
pub const SEPARATOR: char = '\n';

impl Speller {
    /// Parses a dictionary from its `.aff` and `.dic` text (UTF-8).
    ///
    /// # Errors
    /// `SpellerError::Parse` when either text is not a valid Hunspell dictionary.
    pub fn new(aff: &str, dic: &str) -> Result<Self, SpellerError> {
        Dictionary::new(aff, dic)
            .map(|dictionary| Self { dictionary })
            .map_err(|error| SpellerError::Parse(error.to_string()))
    }

    #[must_use]
    pub fn check(&self, word: &str) -> bool {
        self.dictionary.check(word)
    }

    /// Checks `SEPARATOR`-separated words, writing one verdict byte per word (1 correct, 0 not) into
    /// `out`. Returns how many words were checked: all of them, or as many as `out` holds.
    pub fn check_words(&self, words: &str, out: &mut [u8]) -> usize {
        let checker = self.dictionary.checker();
        let mut count = 0;

        for (slot, word) in out.iter_mut().zip(words.split(SEPARATOR)) {
            *slot = u8::from(checker.check(word));
            count += 1;
        }

        count
    }

    /// Corrections for `word`, best first. Without `ngram` only the edit-based pass runs: it answers in
    /// milliseconds where the n-gram pass scans the whole word list.
    #[must_use]
    pub fn suggest(&self, word: &str, ngram: bool) -> Vec<String> {
        let mut out = Vec::new();

        self.dictionary
            .suggester()
            .with_ngram_suggestions(ngram)
            .suggest(word, &mut out);

        out
    }

    /// Accepts `word` from now on, with the casing variants Hunspell derives for a dictionary word. The
    /// word is taken literally: the `.dic` syntax's flag separator is escaped.
    pub fn add(&mut self, word: &str) -> bool {
        // A word with whitespace would be read as a word plus a morphological field.
        if word.is_empty() || word.contains([SEPARATOR, ' ', '\t']) {
            return false;
        }

        // Only "/" is special inside a word (it starts the flags); a backslash before it escapes it.
        self.dictionary.add(&word.replace('/', "\\/")).is_ok()
    }

    /// Writes `items` separated by `SEPARATOR` into one byte buffer (the C ABI's output format).
    #[must_use]
    pub fn join(items: &[String]) -> Vec<u8> {
        let mut out = Vec::with_capacity(items.iter().map(|item| item.len() + 1).sum());

        for (i, item) in items.iter().enumerate() {
            if i > 0 {
                out.push(SEPARATOR as u8);
            }

            out.extend_from_slice(item.as_bytes());
        }

        out
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const AFF: &str = "SET UTF-8\nTRY esianrtolcdugmphbyfvkwzESIANRTOLCDUGMPHBYFVKWZ'\nREP 1\nREP f ph\nSFX S Y 1\nSFX S 0 s .\n";
    const DIC: &str = "4\nhello\nworld/S\ncafé\nphone\n";

    fn speller() -> Speller {
        Speller::new(AFF, DIC).expect("the fixture parses")
    }

    #[test]
    fn checks_stems_and_affixes() {
        let speller = speller();

        assert!(speller.check("hello"));
        assert!(speller.check("worlds"));
        assert!(speller.check("café"));
        assert!(speller.check("Hello"));
        assert!(speller.check("HELLO"));
        assert!(!speller.check("helo"));
        assert!(!speller.check("hellos"));
    }

    #[test]
    fn checks_batches_into_the_output() {
        let speller = speller();
        let mut out = [9u8; 4];

        assert_eq!(
            speller.check_words("hello\nhelo\nworlds\ncafe", &mut out),
            4
        );
        assert_eq!(out, [1, 0, 1, 0]);

        let mut short = [9u8; 2];

        assert_eq!(speller.check_words("hello\nhelo\nworlds", &mut short), 2);
        assert_eq!(short, [1, 0]);
    }

    #[test]
    fn suggests_with_and_without_ngrams() {
        let speller = speller();

        assert_eq!(
            speller.suggest("helo", false).first().map(String::as_str),
            Some("hello")
        );
        assert!(speller.suggest("fone", false).iter().any(|s| s == "phone"));
        assert!(!speller.suggest("wrold", true).is_empty());
        assert!(speller.suggest("hello", false).is_empty() || speller.check("hello"));
    }

    #[test]
    fn adds_words_literally() {
        let mut speller = speller();

        assert!(!speller.check("filen"));
        assert!(speller.add("filen"));
        assert!(speller.check("filen"));
        assert!(speller.check("Filen"));
        assert!(speller.add("A/S"));
        assert!(speller.check("A/S"));
        assert!(!speller.check("A"));
        assert!(speller.add("back\\slash"));
        assert!(speller.check("back\\slash"));
        assert!(!speller.add(""));
        assert!(!speller.add("two words"));
        assert!(!speller.add("line\nbreak"));
    }

    #[test]
    fn rejects_unparsable_dictionaries() {
        assert!(matches!(
            Speller::new("FLAG num\n", "1\nword/abc\n"),
            Err(SpellerError::Parse(_))
        ));
    }

    #[test]
    fn joins_with_the_separator() {
        assert_eq!(Speller::join(&[]), b"");
        assert_eq!(Speller::join(&["a".into(), "bc".into()]), b"a\nbc");
    }

    // The shipped dictionaries load and agree with Hunspell on a few words each. They are built (not
    // committed) by `pnpm run build:dictionaries`; without them this test has nothing to check.
    #[test]
    fn loads_the_shipped_dictionaries() {
        let root = concat!(env!("CARGO_MANIFEST_DIR"), "/../.dictionaries");

        if !std::path::Path::new(&format!("{root}/.complete")).exists() {
            eprintln!("skipped: {root} is not built (pnpm run build:dictionaries)");

            return;
        }
        let cases: &[(&str, &[&str], &[&str])] = &[
            (
                "en_US",
                &["color", "don't", "optimize"],
                &["colour", "optimise", "teh"],
            ),
            ("en_GB", &["colour", "optimise"], &["color", "teh"]),
            ("de_DE", &["Straße", "Donaudampfschifffahrt"], &["Strase"]),
            ("de_CH", &["Strasse"], &["Strase"]),
            ("fr", &["été", "aujourd'hui"], &["etee"]),
            ("da_DK", &["A/S", "hus"], &["huus"]),
        ];

        for (id, good, bad) in cases {
            let aff = std::fs::read_to_string(format!("{root}/{id}/{id}.aff")).expect("aff");
            let dic = std::fs::read_to_string(format!("{root}/{id}/{id}.dic")).expect("dic");
            let speller = Speller::new(&aff, &dic).unwrap_or_else(|error| panic!("{id}: {error}"));

            for word in *good {
                assert!(speller.check(word), "{id} rejects {word}");
            }

            for word in *bad {
                assert!(!speller.check(word), "{id} accepts {word}");
            }
        }
    }
}
