// Conditional formatting rules and their evaluation order.

import {messageOf} from '../../thrown.ts';
import type {Untyped} from '../../untyped.ts';
import {partMapOf} from './package-facts.ts';
import {fixtureBytes, readFixture, readXlsx, Workbook, writeXlsx} from './runtime.ts';
import {attrsOf} from './xml-probes.ts';

export const conditionalFormatting = {
  // Author a conditional-formatting rule, write it, and report the emitted CF XML facts plus what the
  // reader surfaces on reload → { writeOk, writeError, xml:{blockCount, sqrefs, ruleCount, hasDataBar,
  // cfvoCount, hasColor, wellFormed}, reload:{type, color, gradient, cfvo} }.
  authorConditionalFormatting(cf: Untyped) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    // Populate the ref's first column so the rule binds to real cells.
    const rows = Number((cf.ref.match(/(\d+)\s*$/) || [])[1] || 3);
    for (let r = 1; r <= rows; r++) sheet.getCell(`A${r}`).value = r / rows;
    let buffer: Uint8Array;
    try {
      sheet.addConditionalFormatting(cf);
      buffer = writeXlsx(workbook);
    } catch (e) {
      return {
        writeOk: false,
        writeError: messageOf(e),
        xml: null,
        reload: null,
      };
    }
    const xml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const cfBlock = (xml.match(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/) || [
      '',
    ])[0];
    const dataBar = (cfBlock.match(/<dataBar\b[\s\S]*?<\/dataBar>|<dataBar\b[^>]*\/>/) || [''])[0];
    const rule =
      readXlsx(buffer).getWorksheet('S')?.conditionalFormattings?.[0]?.rules?.[0] ?? null;
    return {
      writeOk: true,
      writeError: null,
      xml: {
        blockCount: [...xml.matchAll(/<conditionalFormatting\b/g)].length,
        sqrefs: [...xml.matchAll(/<conditionalFormatting\b[^>]*sqref="([^"]*)"/g)].map((m) => m[1]),
        ruleCount: [...cfBlock.matchAll(/<cfRule\b/g)].length,
        hasDataBar: /<dataBar\b/.test(cfBlock),
        cfvoCount: [...dataBar.matchAll(/<cfvo\b/g)].length,
        hasColor: /<color\b/.test(dataBar),
        wellFormed: cfBlock
          ? !/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(cfBlock)
          : false,
      },
      reload: rule
        ? {
            type: rule.type ?? null,
            color: rule.color ? (rule.color.argb ?? null) : null,
            gradient: rule.gradient ?? null,
            cfvo: (rule.cfvo || []).map((v) => ({type: v.type ?? null, value: v.value ?? null})),
          }
        : null,
    };
  },

  // Apply a stopIfTrue rule, write, reload → { xmlHasStopIfTrue, reloadStopIfTrue }.
  conditionalFormattingStopIfTrue() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('A1').value = 5;
    sheet.addConditionalFormatting({
      ref: 'A1:A10',
      rules: [
        {
          type: 'cellIs',
          operator: 'greaterThan',
          formulae: [3],
          stopIfTrue: true,
          style: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFFF0000'}}},
        },
      ],
    });
    const buffer = writeXlsx(workbook);
    const xml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const rule = readXlsx(buffer).getWorksheet('S')?.conditionalFormattings?.[0]?.rules?.[0];
    return {
      xmlHasStopIfTrue: /stopIfTrue="1"/.test(xml),
      reloadStopIfTrue: rule ? (rule.stopIfTrue ?? false) : null,
    };
  },

  // Read a fixture's first-sheet conditional-formatting facts, write it back, and report the same
  // before/after → { source, rewritten } each { blockCount, rules:[{type, dxfId, priority}] }.
  roundtripFixtureConditionalFormatting(rel: string) {
    const cfFacts = (xml: string) => ({
      blockCount: [...xml.matchAll(/<conditionalFormatting\b/g)].length,
      rules: [...xml.matchAll(/<cfRule\b([^>]*?)\/?>/g)].map((m) => {
        const a = attrsOf(`<x ${m[1]}>`);
        return {type: a.type ?? null, dxfId: a.dxfId ?? null, priority: a.priority ?? null};
      }),
    });
    const srcParts = partMapOf(fixtureBytes(rel));
    const srcName = Object.keys(srcParts).find((n) => n.endsWith('sheet1.xml'));
    const source = cfFacts(srcName === undefined ? '' : (srcParts[srcName] ?? ''));
    const outXml = partMapOf(writeXlsx(readFixture(rel)))['xl/worksheets/sheet1.xml'] || '';
    return {source, rewritten: cfFacts(outXml)};
  },

  // Read a fixture, write it back, and report its first sheet's conditional formats → { read, source,
  // rewritten }. `read` is what the model surfaces, one entry per range: [{ref, rules:[{type,
  // formulae}]}]. `source` and `rewritten` are what each package stores, { classic:[{sqref, rule}],
  // extension:[{sqref, rule, linksAClassicRule}] }, `rule` being the cfRule's markup with the id that
  // ties an extension rule to a classic one blanked, since that id is the writer's to choose. Every
  // list is sorted by range: a rule's place in either form means nothing, its priority does.
  conditionalFormatRulesAsStored(rel: string) {
    const stored = (pkg: Uint8Array) => {
      const parts = partMapOf(pkg);
      const name = Object.keys(parts).find((n) => n.endsWith('sheet1.xml'));
      const xml = name === undefined ? '' : (parts[name] ?? '');
      const linkedIds = new Set([...xml.matchAll(/<x14:id>([^<]*)<\/x14:id>/g)].map((m) => m[1]));
      const blank = (rule: string) =>
        rule.replace(/<x14:id>[^<]*<\/x14:id>/g, '<x14:id/>').replace(/\sid="[^"]*"/g, '');
      const bySqref = (a: {sqref: string; rule: string}, b: {sqref: string; rule: string}) =>
        a.sqref.localeCompare(b.sqref) || a.rule.localeCompare(b.rule);
      const classic = [
        ...xml.matchAll(
          /<conditionalFormatting\b[^>]*\bsqref="([^"]*)"[^>]*>([\s\S]*?)<\/conditionalFormatting>/g,
        ),
      ].flatMap(([, sqref = '', body = '']) =>
        [...body.matchAll(/<cfRule\b[^>]*\/>|<cfRule\b[^>]*>[\s\S]*?<\/cfRule>/g)].map((m) => ({
          sqref,
          rule: blank(m[0]),
        })),
      );
      const extension = [
        ...xml.matchAll(
          /<x14:conditionalFormatting\b[^>]*>([\s\S]*?)<\/x14:conditionalFormatting>/g,
        ),
      ].flatMap(([, body = '']) => {
        const sqref = /<xm:sqref>([^<]*)<\/xm:sqref>/.exec(body)?.[1] ?? '';
        return [
          ...body.matchAll(/<x14:cfRule\b[^>]*\/>|<x14:cfRule\b[^>]*>[\s\S]*?<\/x14:cfRule>/g),
        ].map((m) => ({
          sqref,
          rule: blank(m[0]),
          linksAClassicRule: linkedIds.has(/\sid="([^"]*)"/.exec(m[0])?.[1] ?? ''),
        }));
      });
      return {classic: classic.sort(bySqref), extension: extension.sort(bySqref)};
    };
    const workbook = readFixture(rel);
    const read = (workbook.worksheets[0]?.conditionalFormattings ?? [])
      .map((set) => ({
        ref: set.ref,
        rules: set.rules.map((rule) => ({type: rule.type, formulae: rule.formulae ?? []})),
      }))
      .sort((a, b) => a.ref.localeCompare(b.ref));
    return {read, source: stored(fixtureBytes(rel)), rewritten: stored(writeXlsx(workbook))};
  },

  // Read a fixture and report each data bar on its first sheet as the model surfaces it, sorted by range
  // → [{ref, anchors, color, minLength, maxLength, showValue, gradient, border, borderColor, direction,
  // negativeFillColor, negativeBorderColor, negativeBarColorSameAsPositive,
  // negativeBarBorderColorSameAsPositive, axisPosition, axisColor}]. An anchor is its type, with `:value`
  // when it has one; a colour is its ARGB; a facet the bar does not state is `null`.
  dataBarsAsRead(rel: string) {
    const argb = (color: {argb?: string} | undefined) => color?.argb ?? null;
    return (readFixture(rel).worksheets[0]?.conditionalFormattings ?? [])
      .flatMap((set) =>
        set.rules
          .filter((rule) => rule.type === 'dataBar')
          .map((rule) => ({
            ref: set.ref,
            anchors: (rule.cfvo ?? []).map((anchor) =>
              anchor.value === undefined ? anchor.type : `${anchor.type}:${anchor.value}`,
            ),
            color: argb(rule.color),
            minLength: rule.minLength ?? null,
            maxLength: rule.maxLength ?? null,
            showValue: rule.showValue ?? null,
            gradient: rule.gradient ?? null,
            border: rule.border ?? null,
            borderColor: argb(rule.borderColor),
            direction: rule.direction ?? null,
            negativeFillColor: argb(rule.negativeFillColor),
            negativeBorderColor: argb(rule.negativeBorderColor),
            negativeBarColorSameAsPositive: rule.negativeBarColorSameAsPositive ?? null,
            negativeBarBorderColorSameAsPositive: rule.negativeBarBorderColorSameAsPositive ?? null,
            axisPosition: rule.axisPosition ?? null,
            axisColor: argb(rule.axisColor),
          })),
      )
      .sort((a, b) => a.ref.localeCompare(b.ref));
  },
};
