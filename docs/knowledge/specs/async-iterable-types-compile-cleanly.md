# Async-iterable reader types must compile against a declared minimum

Cluster: types

Reported as "getting cannot find name AsyncGenerator while build".

## Desired behavior

The library's published type declarations must type-check cleanly in a fresh consumer project that
meets the library's _documented_ minimum TypeScript version and `lib`/`target` settings, without the
consumer being forced to enable `skipLibCheck`, hand-edit shipped `.d.ts` files, or add ad-hoc `lib`
entries beyond what the library documents as required.

Streaming readers (workbook, worksheet and row readers) are exposed as async iterables. Their
declared return types must not silently assume the consumer has enabled `esnext.asynciterable` (or
an equally new `lib`) unless that requirement is:

1. declared as a hard minimum in the package's supported-environment docs, and
2. verified by a type-level test in CI so the assumption cannot regress.

## Prior art / evidence

Consumers on older TypeScript, or with `lib: ["es2017", "dom"]`, hit `TS2304: Cannot find name
'AsyncGenerator'` from the shipped `.d.ts`. Community fixes were all consumer-side band-aids:
`skipLibCheck: true`, commenting out the offending declaration lines, or upgrading
TS/`@types/node`/lib target. That is a symptom of the published types depending on consumer lib
flags rather than being self-contained.

## Design stance for ts-xlsx

- Pick and document a firm minimum TypeScript version and required `lib` for consumers.
- Add a type-level (tsd or `expectTypeOf`-style) fixture that compiles the public `.d.ts` under
  exactly that documented minimum config, with no `skipLibCheck`, so a regression is caught in CI.
- Prefer that async-iteration types resolve from the library's own declared `lib` requirement
  rather than leaking an undeclared newer-lib dependency onto consumers.

## Open questions

- What is the committed minimum TypeScript version and `lib`/`target` for consumers of ts-xlsx?
- Should the async-iterator API be typed via `AsyncIterableIterator`/`AsyncGenerator` (which needs
  the async-iterable lib) or a narrower hand-rolled interface to reduce the lib requirement?
