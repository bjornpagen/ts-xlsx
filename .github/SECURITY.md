# Security policy

A spreadsheet is an attacker-controlled zip of attacker-controlled XML. Every parser path in
this library is treated as hostile-input-facing (`CLAUDE.md` §3), so a file that makes it
allocate without bound, spin without end, or write outside where it was told to is a
vulnerability here — not a malformed-input curiosity.

## Reporting

**Report privately: <https://github.com/shbernal/ts-xlsx/security/advisories/new>.**

Do not open a public issue, and do not demonstrate the bug in a public pull request. The
advisory thread is private until a fix ships, and it can mint a CVE and credit you at
publication.

Include the reproduction the way the [agent report
form](ISSUE_TEMPLATE/agent-report.yml) asks for it — a self-contained script that _builds_
its malicious input rather than attaching one. A proof-of-concept we can regenerate becomes a
permanent regression case; an attached binary usually cannot be committed.

**Never attach a workbook containing real data**, even privately. If the finding only
reproduces against a real file, describe the structural feature you suspect and we will
synthesize one.

### What to expect

This project is maintained by one person and a fleet of agents. There is no SLA and no
bounty. What there is: an acknowledgement within a week, a fix on the current release line
rather than a backport queue, and an advisory published with your name on it unless you ask
otherwise.

## Supported versions

Fixes ship on the latest published version. There are no backports and no long-term support
branches — major bumps are cheap here by design (`CLAUDE.md` §1.3), so the supported upgrade
path is forward.

## In scope

- Resource exhaustion from a crafted file: zip bombs, decompression ratios, entity expansion,
  quadratic parsing, unbounded buffering of a stream that never terminates.
- Anything that escapes the archive during extraction — path traversal, absolute entry names,
  symlink entries.
- XML external entities, remote references, or any parse that reaches the network or the
  filesystem on its own.
- Prototype pollution reachable from parsed content (a sheet name, a defined name, a part
  name, a relationship target).
- Type or API surfaces that make the unsafe call the easy one — a documented escaping or
  validation helper that does not actually escape or validate.

## Out of scope

- Formula injection into cells _you_ write. If you place untrusted text in a cell and Excel
  later evaluates it as a formula, that is your application's trust boundary; this library
  writes what you tell it to. A documented escape hatch that fails to escape _is_ in scope.
- Vulnerabilities in Excel, LibreOffice, or any viewer that opens what we emit — report those
  to their vendors. We do want to know if we can be made to emit the trigger.
- Anything requiring an attacker who already executes code in your process.
- A crash, on its own, when it is a typed error thrown deliberately at a validated boundary.
  Report it as a bug; it becomes a security matter when it is unbounded work or memory rather
  than a refusal.
