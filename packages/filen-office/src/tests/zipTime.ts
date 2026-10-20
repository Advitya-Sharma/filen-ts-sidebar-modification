// The entry time for fflate's zipSync in test fixtures (its default is the clock): a save keeps an entry's time,
// so two builds of one fixture, however far apart, must be the same bytes. Midday UTC, a date in fflate's range
// (1980-2099, local time) in every time zone a test sets.
export const ZIP_ENTRY_TIME = Date.UTC(2000, 0, 1, 12)
