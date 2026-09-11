# Excel decodes `_xHHHH_` in cell text, in a threaded comment, and in a print header, in one left-to-right pass

Cluster: xlsx-io

## Scenario

XML 1.0 has no syntax for a C0 control character. `&#1;` is not an escape for U+0001; it is another
spelling of the same illegal document. SpreadsheetML's answer is an in-band convention, `_xHHHH_`,
where `HHHH` is four hex digits naming a code point. A writer that wants to carry such a character
in a cell has to use it, and a reader that wants to read Excel's files has to undo it.

That much is widely implemented (POI and openpyxl both do it) and widely repeated. It is still a
claim about what an _application_ does, not what a schema says, and `ST_Xstring` types a sheet name
exactly as it types a `<t>`, so the schema cannot even tell you where the convention applies. The
questions that actually decide an implementation are narrower than "does Excel decode it":

- Does it decode in the shared-strings pool, in an inline `<is>`, or both?
- Does it decode in the cached `<v>` of a `t="str"` formula cell, which is a value but not a `<t>`?
- What happens to `_x005F_x0041_`, the escaped underscore in front of an escape? One pass yields the
  literal text `_x0041_`; two passes collapse it to `A`. Both are defensible from the prose.
- Is the decode restricted to characters XML could not have carried anyway?
- How strict is the grammar? Does `_x041_` or `_xZZZZ_` decode to anything?
- Does it reach beyond cell text at all? A threaded comment's `<text>` is prose a human typed, but it
  is a different element in a Microsoft extension namespace that documents no escape of its own. A
  print header's `<oddHeader>` is prose too, but its text already carries an in-band syntax of its own,
  the `&`-prefixed section and format codes, which is exactly the kind of thing that makes an
  application special-case an element.

## Measured behavior

Excel Desktop (Microsoft 365, version 16.0, build 20228, Windows), driven headless over COM with
`Calculation` forced to manual so a recalculation could not overwrite a cached `<v>` before it was
read. A package written by this library, then textually patched so the escape sequences reach the
parts verbatim rather than being escaped by our own writer; validated clean by `ooxml-validate`
before opening. Values read through `Range.Value2` and reported as character codes, so nothing is
lost to console rendering.

| Cell | On disk                                        | Read back                |
| ---- | ---------------------------------------------- | ------------------------ |
| A1   | `<c t="inlineStr"><is><t>_x0041_</t></is></c>` | `A`                      |
| A7   | `t="s"` into `<si><t>_x0041_</t></si>`         | `A`                      |
| B1   | `<c t="str"><f>A1</f><v>_x0041_</v></c>`       | `A`                      |
| A2   | `<t>_x005F_x0041_</t>`                         | `_x0041_` (7 characters) |
| A3   | `<t>_xZZZZ_</t>`                               | `_xZZZZ_`                |
| A4   | `<t>_x041_</t>`                                | `_x041_`                 |
| A5   | `<t>_x00041_</t>`                              | `_x00041_`               |
| A6   | `<t>a_x0009_b</t>`                             | `a`, U+0009, `b`         |

Five facts fall out.

**The pool and the inline string behave identically.** A1 and A7 are the same escape reaching the
cell by different routes, and both decode. So the convention belongs to the `<t>`, not to the part
it happens to live in.

**The cached result of a string formula decodes too.** B1 is the case the schema cannot settle and
the one most likely to be assumed wrong in either direction. It decodes, so a `t="str"` `<v>` is
cell text for this purpose despite not being a `<t>`.

**The decode is one left-to-right pass.** A2 is the proof: `_x005F_` yields `_`, scanning resumes
past it on `x0041_`, and with no leading underscore left that is ordinary text. Excel reads the cell
as the literal `_x0041_`. A decoder that resolves `_x005F_` in a pass of its own, before or after
the rest, reads that cell as `A` and silently destroys every value that legitimately looks like an
escape.

**The grammar is closed and exact.** Four hex digits, no more and no fewer, between `_x` and `_`.
A3, A4 and A5 are returned untouched, so a reader must not be liberal here: near-misses are ordinary
text and stay that way.

**The decode is unconditional, not a repair.** A6 carries a tab, which XML 1.0 permits literally, and
Excel decodes it anyway. A reader that only undid the escape for characters XML cannot represent
would disagree with Excel about files Excel itself wrote.

## The same convention in a threaded comment's `<text>`

The measurement above says the escape works in _cell text_. It says nothing about the modern threaded
comments, whose body is a `<text>` in
`http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments`: a different element, in a
Microsoft extension namespace, with no documented escape. That asymmetry was visible to callers: a
legacy note's body is a `<t>` in a `CT_Rst` and carried a control character happily, while the same
string in a threaded comment was refused. Same probe, same question, second element (2026-08-25).

The probe package this library wrote carries one single-message thread per cell, and was patched so
each body reached the part verbatim, validated clean by `ooxml-validate`, then opened headless and
read through `Range.CommentThreaded.Text`, reported as character codes. The legacy fallback
`<comment>` beside each thread was deliberately left holding the _unpatched_ placeholder, so a
readback that reports the patched text proves Excel read the 2018 part and not the fallback.

| Cell | `<text>` on disk | Read back                |
| ---- | ---------------- | ------------------------ |
| A1   | `_x0041_`        | `A`                      |
| A2   | `_x005F_x0041_`  | `_x0041_` (7 characters) |
| A3   | `_xZZZZ_`        | `_xZZZZ_`                |
| A4   | `_x041_`         | `_x041_`                 |
| A5   | `_x00041_`       | `_x00041_`               |
| A6   | `a_x0009_b`      | `a`, U+0009, `b`         |
| A7   | `plain`          | `plain`                  |
| A8   | `_x0001_`        | U+0001                   |

**It is the same convention, cell for cell.** Every row matches the cell-text table: the escape
decodes, the near-misses do not, the escaped underscore resolves in one left-to-right pass, and a
character XML could have carried is decoded anyway. Nothing about the extension namespace changes the
grammar.

**Excel writes it back, so this is its representation and not merely a tolerated read.** Re-saving the
same package through Excel (`SaveAs`, `xlOpenXMLWorkbook`) and reading the emitted part shows A8 back
on disk as `_x0001_` and A2 re-escaped to `_x005F_x0041_`. That is Excel's own writer performing the
same underscore-first escape this library performs. A1 comes back as the plain `A` it decoded to, and
A6 as a literal tab rather than `_x0009_`, so Excel escapes only what it must here.

That settles the open question the follow-up list carried: the threaded comment's `<text>` joins the
escaping group.

## The same convention in a print header's `<oddHeader>`

A print header was the last element still guessed rather than measured, and guessed the other way: its
text was refused on write when it held an unrepresentable character, and read back verbatim. That was
consistent, but it was decided from the format's shape rather than measured. The `&`-prefixed section
and format codes look like the only in-band syntax a header has, so a second convention layered on top
seemed unlikely. The threaded comment had already shown that shape is a poor guide.

Same probe, same eight rows, third element (2026-08-25). The package this library wrote carries one
`<headerFooter>` whose sections were patched so each escape reached the part verbatim; it validated
clean by `ooxml-validate`, was opened headless, and every section was read through `PageSetup`
(`LeftHeader`/`CenterHeader`/`RightHeader`, the footer trio, and `EvenPage.*.Text`), reported as
character codes.

| Slot                | Section on disk | Read back                |
| ------------------- | --------------- | ------------------------ |
| `<oddHeader>` `&L`  | `_x0041_`       | `A`                      |
| `<oddHeader>` `&C`  | `_x005F_x0041_` | `_x0041_` (7 characters) |
| `<oddHeader>` `&R`  | `_xZZZZ_`       | `_xZZZZ_`                |
| `<oddFooter>` `&L`  | `_x041_`        | `_x041_`                 |
| `<oddFooter>` `&C`  | `_x00041_`      | `_x00041_`               |
| `<oddFooter>` `&R`  | `a_x0009_b`     | `a`, U+0009, `b`         |
| `<evenHeader>` `&L` | `plain`         | `plain`                  |
| `<evenFooter>` `&L` | `_x0001_`       | U+0001                   |

**It is the same convention again, row for row.** The escape decodes, the near-misses do not, the
escaped underscore resolves in one left-to-right pass, and a character XML could have carried is
decoded anyway. The `&`-prefixed codes are untouched by it: they are ordinary text to the decoder,
which is why the two syntaxes can share one string without either needing to know about the other.

**Excel writes it back here too.** Re-saving through Excel (`SaveAs`, `xlOpenXMLWorkbook`) emits
`&amp;L_x0001_` for the even footer and `&amp;C_x005F_x0041_` for the literal, while the decoded `A`
comes back plain and the tab comes back literal. So the escape is Excel's own representation for a
header, not a tolerated read.

The consequence for this library is a refusal removed, not merely a decode added: a header may now
carry any character a note or a cell may.

## What follows for this library

Both directions are in `src/xml/`: `escapeSpreadsheetText` in `xml.ts`, `decodeSpreadsheetText` in
`xml-scan.ts`. They are inverses, and three corpus cases lock that with the packages above as their
fixtures: `escaped-characters-in-cell-text-decode-on-read` for cell text,
`threaded-comment-text-decodes-xhhhh-escape` for the conversation body, and
`print-header-text-decodes-xhhhh-escape` for the header/footer definition.

Where the convention applies is a separate decision the writer makes, and no measurement settles it in
general: they say the escape _works_ in these three places, not that it works anywhere else. This
library escapes what a human typed (cell text, a note body, a threaded message, a print header) and
refuses everything structural. The module header of `src/xml/xml.ts` states the line and why it falls
there. The pattern across three elements is worth stating, though, because it is now the prior for the
fourth: every element measured so far that carries _prose a person typed_ decodes, whatever namespace
it lives in and whatever other syntax its text already carries.

## Provenance

`source: excel-desktop-verification`, the tier ADR-0013 describes. One Excel build, one host.

The COM tier is sufficient here: every fact above is state Excel reports directly, through
`Range.Value2` for a cell, `Range.CommentThreaded.Text` for a conversation, and `PageSetup` for a
header, so nothing needed the rendering tier. Reproducible from the repo: all three probe packages are
committed as corpus fixtures,
`test/corpus/fixtures/escaped-characters-in-cell-text-decode-on-read/xhhhh-escapes.xlsx` (open it and
read A1:A7 and B1),
`test/corpus/fixtures/threaded-comment-text-decodes-xhhhh-escape/xhhhh-escapes-in-threads.xlsx` (open
it and read the threaded comment on A1:A8), and
`test/corpus/fixtures/print-header-text-decodes-xhhhh-escape/xhhhh-escapes-in-headers.xlsx` (open it
and read the eight header/footer sections through File → Page Setup, or over COM).

Related: `excel-repair-on-open-structural-constraints`.
