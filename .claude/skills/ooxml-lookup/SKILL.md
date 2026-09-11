---
name: ooxml-lookup
description: >-
  Query the ECMA-376 (Office Open XML) schema offline to find out what is legal
  in a .docx, .xlsx or .pptx file. Use when writing, reading or debugging OOXML
  markup by hand: what may go inside an element and in what order, what
  attributes it takes, what values those attributes accept, which namespace or
  prefix to write, and what the difference is between the Transitional and
  Strict profiles. Also resolves a schema validation error into "here is what
  would have been legal at that position". Covers wordprocessingml,
  spreadsheetml, presentationml, drawingml and vml. Does NOT validate files
  (that is ooxml-validate), does not generate documents, and has no
  specification prose, behaviour notes or semantic search — it answers from the
  XSD schema graph only.
---

# ooxml-lookup

The ECMA-376 schema as a local SQLite graph, queried through one CLI. No
network, no account, no install — `node scripts/ooxml.mjs` from a bare checkout.

Every command prints JSON on stdout and nothing else.

## The loop

Most questions are one of these three, in this order.

**1. What is this thing?**

```bash
node scripts/ooxml.mjs element w:tblPr
```

Returns every declaration of the name. More than one is normal and not an
error — OOXML declares element names _locally_, inside types, so `w:tblPr` is
`CT_TblPr` inside `w:CT_Tbl` and `CT_TblPrBase` in three other places. The
`scope.declared_in` field tells you which is which.

**2. What is legal here?**

```bash
node scripts/ooxml.mjs children w:CT_PPr      # what may go inside, in order
node scripts/ooxml.mjs attributes w:CT_Ind    # what attributes it takes
node scripts/ooxml.mjs values s:ST_TwipsMeasure   # what values are legal
```

`children` returns two views of the same thing: `tree`, the nested
sequence/choice/all structure with cardinalities, and `order`, a flat list in
schema order. Read `order` for "what goes here"; read `tree` when the
choice/sequence distinction matters.

`min`/`max` of `-1` means unbounded.

**3. What does this validation error mean?**

```bash
node scripts/ooxml.mjs explain '{"id":"Sch_UndeclaredAttribute",
  "description":"The '\''bogus'\'' attribute is not declared.",
  "xpath":"/w:document[1]/w:body[1]/w:p[1]/w:pPr[1]/w:ind[1]"}'
```

Takes a diagnostic from an `ooxml-validate` report — or the whole report, it
will use the first diagnostic — and answers the question that always follows:
_then what would have been legal there._ The xpath does real work: it
disambiguates a name with several content models by looking at its ancestors.

An id it does not recognise is not a failure; it still tells you what the
schema allows at that position.

## Worked example

You are hand-writing a paragraph and want to indent it.

```bash
$ node scripts/ooxml.mjs attributes w:CT_Ind --compact
{"query":"w:CT_Ind","profile":"transitional","found":true,"type":"w:CT_Ind",
 "count":12,"attributes":[{"name":"end","qualified":true,"use":"optional",
 "type":{"qname":"w:ST_SignedTwipsMeasure",...
```

`firstLine` is typed `s:ST_TwipsMeasure`. What can you write in it?

```bash
$ node scripts/ooxml.mjs values s:ST_TwipsMeasure --compact
{"type":"s:ST_TwipsMeasure","one_of":[
  {"type":"s:ST_UnsignedDecimalNumber","base":"xsd:unsignedLong"},
  {"type":"s:ST_PositiveUniversalMeasure","base":"s:ST_UniversalMeasure",
   "facets":{"pattern":"[0-9]+(\\.[0-9]+)?(mm|cm|in|pt|pc|pi)"}}]}
```

So `w:firstLine="720"` (twips) or `w:firstLine="0.5in"`, and now you know the
unit suffixes are a closed set.

## Things that will bite you

- **Always check the profile.** Answers default to **Transitional**, which is
  what Word, Excel and PowerPoint actually write. Strict uses _different
  namespace URIs for the same vocabulary_ and drops VML entirely. If a document
  declares `purl.oclc.org/ooxml/...` namespaces, pass `--profile strict`.
  `diff` shows what changes for one symbol.
- **A name is not an identity.** `ST_Percentage` is a pattern-restricted string
  in `shared-commonSimpleTypes` and a _union_ in `dml-main`; `ST_Direction` is
  `ltr|rtl` in wml, `horz|vert` in pml and `norm|rev` in dml-diagram. A bare
  name like that comes back as `ambiguous:true` with a `variants` array rather
  than one arbitrary answer, so **check for `variants` before reading a
  result**. Qualify names when you can: `s:ST_Percentage`, not `ST_Percentage`.
- **An empty answer is often correct.** `attributes w:CT_Tbl` returns nothing
  because wml carries table properties as child elements. `found:true` with
  `count:0` means "none", not "lookup failed".
- **`search` is substring, not semantic.** It finds `tblPr`; it will not find
  "how do I make a table border". There are no embeddings here by design.

## Names

Write them any way you like — `w:tblPr`, `wml:tblPr`,
`{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblPr`, or bare
`tblPr`. A bare name that matches several vocabularies returns all of them
rather than guessing.

Prefixes are the conventional ones (`w`, `a`, `p`, `s`, `m`, `r`, `v`, `x`, `c`,
`o`, `xdr`, `wp`, …). Run `namespace <uri>` to go the other way, or
`namespace <prefix>` to see which vocabulary a prefix reaches.

Two of those — `x` for spreadsheetml and `c` for charts — are accepted because
the ecosystem writes them, not because a schema binds them; `namespace` reports
those under `aliases` with a citation, and leaves `prefix` null, which is what
the standard actually says. **Answers come back in the canonical spelling**, so
`element x:worksheet` replies `sml:worksheet`. That is deliberate: `x` is _also_
VML's excel namespace, and printing both as `x:` would render two different
namespaces identically. The lookup checks both and returns whichever actually
has the name.

An ambiguous name returns `ambiguous:true` and a `variants` array, one entry per
distinct meaning. The `message` says which kind of ambiguity it is: several
vocabularies sharing a local name (qualify it, and you get one answer), or one
name declared in several places within a vocabulary — `w:tblPr` — where the
answer depends on the parent and each variant carries `applies_when_declared_in`.

## Direct SQL, when the subcommands do not fit

The database ships with the skill and the connection is read-only:

```bash
node scripts/ooxml.mjs sql "SELECT local_name FROM symbols
  WHERE kind='complexType' AND local_name LIKE 'CT_Tbl%'"
```

Tables: `profiles`, `vocabularies`, `namespaces`, `prefix_aliases`, `symbols`,
`symbol_profiles`, `compositors`, `child_edges`, `group_edges`, `attr_edges`,
`inheritance_edges`, `enums`, `simple_type_facets`, `union_members`.

Two things to know before writing a join. Symbols are keyed on the
**vocabulary**, not the namespace URI, because the two profiles are the same
vocabulary under different URIs — so `symbols` has no profile column and every
_edge_ table does. And `symbols.parent_symbol_id` is `0` for a global
declaration, not NULL.

Prefer the subcommands where they fit: they resolve inheritance and expand
group references, and a raw `child_edges` query will not.

## What this does not do

- **Validate a file.** That is `ooxml-validate`, a separate tool. This one
  answers what the schema permits; it never opens your document.
- **Explain what Word actually does.** The schema is the standard, and
  implementations diverge from it. There are no behaviour notes here.
- **Search the specification text.** No prose, no PDFs, no embeddings. If you
  need to read the specification itself, <https://ooxml.dev> does that well and
  this deliberately does not compete with it.
