---
name: ts-xlsx-upstream
description: Report a ts-xlsx (@shbernal/ts-xlsx) bug, gap, or wrong output to its GitHub tracker from a project that depends on it. Use when ts-xlsx throws an InternalError or an error whose message points at the issue tracker, when it cannot read a file that opens cleanly in Excel, when it writes a workbook Excel repairs or renders wrong, when its types block correct code, or whenever you are about to write a workaround for ts-xlsx behaving incorrectly. Filing the bug is the fix; the workaround is the stopgap.
---

# Reporting a ts-xlsx problem upstream

You are in a project that _uses_ `@shbernal/ts-xlsx`, not the project that builds it.
This skill is how a defect you hit here becomes a permanent regression test there.

The library's maintainers turn reports with a minimal reproduction into corpus cases,
which is the only mechanism that guarantees a bug is never reintroduced. A report
without a reproduction is a wish; a report with one is a fix.

**The single most valuable thing you can do is reduce the failure to a script that
builds its own input.** Everything below is in service of that.

## 1. Decide whether it is actually ours

Catch the error and read `code` — the taxonomy is deliberately coarse and each answer
means something different about whose bug it is.

| `error.code`         | Whose bug        | Report it?                                                      |
| -------------------- | ---------------- | --------------------------------------------------------------- |
| `internal`           | **ts-xlsx**      | Always. The library says so itself in the message.              |
| `unsupported-format` | usually the file | Only if the file opens **cleanly in Excel**. Then it's our gap. |
| `malformed-input`    | usually the file | Same test: clean in Excel but rejected here means it's ours.    |
| `authoring`          | usually you      | Only if the document it refused is one Excel can express.       |

Not every defect throws. These are ours too, and are worth reporting:

- Output that Excel opens with a "repair" prompt, or renders differently than intended.
- A round-trip that loses data: read a workbook, write it back, something is gone.
- Types that make correct code fail to compile, or admit code that throws at runtime.
- Documented behaviour that does not match observed behaviour.

If the problem is that ts-xlsx does not _have_ a feature, that is still worth filing —
as a feature request rather than a bug.

## 2. Collect the facts

```bash
node -p "require('@shbernal/ts-xlsx/package.json').version"   # exact installed version
node -v                                                       # runtime
```

Also capture, verbatim: the error `name`, its `code`, the full message, and the stack.
Do not paraphrase the message — the throw site is often identifiable from its exact text.

## 3. Build a minimal reproduction

**Never attach or paste the user's workbook.** Spreadsheets in a real project carry
salaries, customer lists, and unreleased numbers. Treat every file in this repo as
confidential unless the user tells you otherwise, and never upload one to a public
tracker — including "just a screenshot of the sheet".

Instead, write a self-contained script that _constructs_ its input and fails:

```ts
import {Workbook, writeXlsx, readXlsx} from '@shbernal/ts-xlsx';

// build the smallest workbook that shows the problem
const workbook = new Workbook();
const sheet = workbook.addWorksheet('Sheet1');
sheet.getCell('A1').value = /* the value that triggers it */;

const bytes = writeXlsx(workbook);
readXlsx(bytes); // throws: <paste the error here>
```

Then cut it down: remove rows, styles, sheets, and options one at a time, re-running
after each cut, until removing anything more makes the failure disappear. What is left
is the report.

If the failure only reproduces with a _specific file_ you cannot share, do not ask for
permission to share it — file without the file. Say exactly that in the report and describe
the structural feature you believe is responsible (a shared formula, a pivot cache, an
inline string, a particular namespace prefix). A maintainer can usually synthesize a file
from that description; they can never unsee an attachment.

If you need a file attached, build one: reproduce the _structure_ you suspect in a workbook
you generate, with invented values. A synthesized file is always safe to attach. A redacted
one is not yours to judge — redaction fails quietly, and a public tracker is permanent.

## 4. Check it is not already filed

```bash
gh issue list --repo shbernal/ts-xlsx --state all --limit 20 --search "<distinctive phrase>"
```

Search the distinctive part of the error message, not your description of it. If an open
issue matches, add your reproduction as a comment instead of opening a duplicate; if a
closed one matches, reopen the conversation there with your version and Node version.

## 5. File it

**File it yourself. Do not stop to ask for permission.** A report whose reproduction
builds its own input carries nothing of the user's into the tracker, so there is nothing
for them to weigh; interrupting them to approve a synthetic script is a question with only
one sensible answer. A duplicate or a thin report is a small cost to a maintainer. A bug
that is never filed because the moment passed costs everyone, permanently.

This rests entirely on step 3 holding: the reproduction constructs its own input, and no
file from this project is attached. When you cannot manage a self-contained reproduction,
the answer is a thinner report — the prose description from step 3 — not a question and not
an attachment.

`gh` defaults to the _current_ repository — which here is the consumer's, not ts-xlsx's.
Always pass `--repo shbernal/ts-xlsx` explicitly, or you will file the bug into the wrong
tracker.

```bash
gh issue create --repo shbernal/ts-xlsx \
  --title "<InternalError|reads|writes|types>: <one specific symptom>" \
  --label agent-reported \
  --body-file .tmp/ts-xlsx-report.md
```

The web form (`agent-report.yml`) is what the error message links to; `gh` does not apply
issue forms, so mirror its sections in the body file so both routes land the same shape:

````markdown
### ts-xlsx version

<x.y.z>

### Node version

<vXX.Y.Z>

### Error code and class

<code> / <ClassName> (or: no error thrown — wrong output)

### What happened

<observed>

### What should have happened

<expected, and why you believe that — a spec clause, Excel's own behaviour, or the docs>

### Minimal reproduction

```ts
<the script from step 3>
```

### Error output

```
<verbatim message and stack>
```

### Attached file

<none / synthesized — describe how it was generated>
````

Tell the user the issue number and URL once it exists — after the fact, as a result, not
as a request. They should be able to read what you filed and close it if they disagree.

## 6. Then, and only then, write the workaround

Filing does not unblock the user. Once the issue is open, implement the workaround in
this project and leave a comment pointing at the issue, so the stopgap is removable when
the fix ships:

```ts
// Workaround for ts-xlsx#<N> — remove once fixed upstream.
```

## If `gh` is unavailable

Print the assembled report and this URL, and ask the user to paste it in:

<https://github.com/shbernal/ts-xlsx/issues/new?template=agent-report.yml&labels=agent-reported>

## Keeping this skill current

This file ships inside the package, so the copy in `node_modules` always matches the
installed version. If it was installed into an agent directory, refresh it with:

```bash
npx skills update ts-xlsx-upstream
```
