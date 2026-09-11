// A relationship type as ISO/IEC 29500 Strict spells it, read in the Transitional spelling.
//
// Strict is the same vocabulary as Transitional under other names. For a relationship that is one
// namespace, `http://purl.oclc.org/ooxml/officeDocument/relationships/` for
// `http://schemas.openxmlformats.org/officeDocument/2006/relationships/`, and two renamed segments.
// Every write is a Transitional package, and a relationship the reader preserves is written back under
// the Type it was read with, so a Strict one landed in a Transitional package spelled in the other
// profile's names. Translating as the container reads a `.rels` part means nothing past it ever meets
// a Strict Type: the reader recognises a Strict package's relationships by the names it already knows,
// and a preserved one is written in the spelling of the package it lands in. The parts themselves are
// translated by the XML reader that captures them (`io/xlsx/strict-parts.ts`), since a binary workbook
// has no Strict form.

const STRICT_RELATIONSHIPS = 'http://purl.oclc.org/ooxml/officeDocument/relationships/';
const TRANSITIONAL_RELATIONSHIPS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

// The two relationship types whose last segment Strict renamed.
const RENAMED: ReadonlyMap<string, string> = new Map([
  ['extendedProperties', 'extended-properties'],
  ['customProperties', 'custom-properties'],
]);

/** A relationship type in its Transitional spelling: a Strict one translated, any other as it is. */
export function transitionalRelationshipType(type: string): string {
  if (!type.startsWith(STRICT_RELATIONSHIPS)) return type;
  const name = type.slice(STRICT_RELATIONSHIPS.length);
  return `${TRANSITIONAL_RELATIONSHIPS}${RENAMED.get(name) ?? name}`;
}
