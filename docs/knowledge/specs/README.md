# Behaviour spec notes

One file per behaviour, hand-authored and implementation-blind. Each was distilled from
the upstream ExcelJS backlog during the harvest (ADR-0007): a real scenario someone hit, a
statement of what the library should do about it, and the questions that were still open
when the note was written. They describe _behaviour_, never a design, which is why they
outlived the codebase they were extracted from.

A note is not a corpus case. A case is executable and pins the library down; a note is the
evidence a case was written from, and a note whose questions are still open is work that
has not been done. Read one to learn what Excel actually does in a corner; read
[`test/corpus/`](../../../test/corpus/) to learn what this library is held to.

Most notes carry the same three headings:

| Heading           | What it holds                                                      |
| ----------------- | ------------------------------------------------------------------ |
| Scenario          | The situation as the person who reported it met it                 |
| Desired behaviour | What should happen, stated without reference to any implementation |
| Open questions    | What was undecided when the note was written                       |

A `Cluster:` line near the top groups a note with its neighbours (`streaming`, `styles`,
`types`, `images` and about thirty more). It is a reading aid, not a schema.

When a decision needs grounding in the format itself rather than in a report, these notes
are the third tier, after the vendored ECMA-376 schema graph and the Microsoft Learn
search: see [the agent correctness playbook](../../agent-correctness-playbook.md).
