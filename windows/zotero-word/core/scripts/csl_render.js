#!/usr/bin/env node
/* Read-only CSL expected-output renderer. No engine is vendored here.
 *
 * The caller must bind processor_js and abbreviation_js to the trusted static
 * Zotero distribution, and bind styles, locales, abbreviation data, item
 * identities and citation order to the current run. A vm is not a security
 * boundary for arbitrary supplied JavaScript. Never use this to create fields.
 */
'use strict';

const fs = require('fs');
const vm = require('vm');

function requireValue(ok, message) {
  if (!ok) throw new Error(message);
}

function render(input) {
  requireValue(input && typeof input === 'object' && !Array.isArray(input), 'Expected an input object');
  const keys = new Set(['processor_js', 'abbreviation_js', 'abbreviations', 'locales', 'style_xml', 'style_id',
    'automatic_journal_abbreviations', 'itemsByID', 'citations', 'locale']);
  requireValue(Object.keys(input).every(key => keys.has(key)), 'Unknown expected-render input key');
  for (const key of ['processor_js', 'style_xml', 'style_id']) {
    requireValue(typeof input[key] === 'string' && input[key].trim(), 'Missing ' + key);
  }
  requireValue(input.processor_js.length <= 2 * 1024 * 1024 && input.style_xml.length <= 2 * 1024 * 1024,
    'Processor or style exceeds bounded size');
  requireValue(!/<!DOCTYPE|<!ENTITY/i.test(input.style_xml), 'DOCTYPE/entity style declarations are unsupported');
  requireValue(typeof input.automatic_journal_abbreviations === 'boolean', 'Abbreviation option must be explicit');
  requireValue(input.locales && typeof input.locales === 'object' && !Array.isArray(input.locales), 'Missing bound locale XML');
  for (const [name, xml] of Object.entries(input.locales)) {
    requireValue(/^[A-Za-z]{2,3}(?:-[A-Za-z0-9]{2,8})*$/.test(name) && typeof xml === 'string'
      && xml.length <= 1024 * 1024 && !/<!DOCTYPE|<!ENTITY/i.test(xml), 'Invalid bound locale ' + name);
  }
  requireValue(input.locale === undefined || (typeof input.locale === 'string' && input.locale), 'Invalid optional locale');
  requireValue(input.itemsByID && typeof input.itemsByID === 'object' && !Array.isArray(input.itemsByID), 'Missing item map');
  requireValue(Array.isArray(input.citations) && input.citations.length > 0 && input.citations.length <= 20000,
    'Expected a nonempty bounded citation sequence');
  const ids = new Set();
  for (const citation of input.citations) {
    requireValue(citation && typeof citation === 'object' && typeof citation.citationID === 'string'
      && citation.citationID && !ids.has(citation.citationID), 'Missing or duplicate citation identity');
    ids.add(citation.citationID);
    requireValue(citation.properties && typeof citation.properties === 'object'
      && Number.isInteger(citation.properties.noteIndex) && citation.properties.noteIndex >= 0,
      'Citation needs an observed nonnegative noteIndex');
    requireValue(Array.isArray(citation.citationItems) && citation.citationItems.length > 0,
      'Empty native citation cluster is unsupported');
    for (const item of citation.citationItems) {
      requireValue(item && ['string', 'number'].includes(typeof item.id)
        && Object.hasOwn(input.itemsByID, String(item.id)), 'Unbound citation item identity');
      const source = input.itemsByID[String(item.id)];
      requireValue(source && typeof source === 'object' && String(source.id) === String(item.id), 'Item map ID differs from citation ID');
    }
  }
  if (input.automatic_journal_abbreviations) {
    requireValue(typeof input.abbreviation_js === 'string'
      && input.abbreviation_js.startsWith('Zotero.Cite.getAbbreviation = new function () {')
      && input.abbreviation_js.trimEnd().endsWith('}();') && input.abbreviation_js.length <= 65536,
      'Missing bounded static Zotero abbreviation handler');
    requireValue(input.abbreviations && typeof input.abbreviations === 'object'
      && input.abbreviations.info && input.abbreviations.info.name && input.abbreviations.info.URI,
      'Missing bound abbreviation database');
  }

  // No process, require, filesystem, networking or live Zotero APIs are exposed
  // to the processor. The only data entering this context is the bound snapshot.
  const context = vm.createContext({ input: JSON.parse(JSON.stringify(input)) });
  vm.runInContext(input.processor_js, context, { timeout: 15000, filename: 'bound-zotero-citeproc.js' });
  requireValue(context.CSL && typeof context.CSL.Engine === 'function', 'Bound source did not define CSL.Engine');
  vm.runInContext(`
    var rendererWarnings = [];
    var abbreviationCategoriesWithoutData = new Set();
    CSL.debug = function () { rendererWarnings.push(Array.from(arguments).map(String).join(' ')); };
    var abbreviationCategories = new Set();
    if (input.automatic_journal_abbreviations) {
      for (var jurisdiction of Object.keys(input.abbreviations)) {
        var categories = input.abbreviations[jurisdiction];
        if (categories && typeof categories === 'object') {
          for (var category of Object.keys(categories)) abbreviationCategories.add(category);
        }
      }
      var Zotero = {
        Cite: {}, CiteProc: { CSL: CSL }, DataDirectory: { dir: 'verified-memory-snapshot' },
        File: {
          pathToFile: function () { return {
            append: function (name) { if (name !== 'abbreviations.json') throw new Error('Unexpected abbreviation file'); },
            exists: function () { return false; }
          }; },
          getContentsFromURL: function (url) {
            if (url !== 'resource://zotero/schema/abbreviations.json') throw new Error('Unexpected abbreviation resource');
            return JSON.stringify(input.abbreviations);
          }
        },
        Exception: { Alert: class extends Error { constructor() { super(Array.from(arguments).map(String).join(' ')); } } },
        debug: function () {}
      };
    }
  `, context, { timeout: 15000, filename: 'read-only-render-context.js' });
  if (input.automatic_journal_abbreviations) {
    vm.runInContext(input.abbreviation_js, context, { timeout: 15000, filename: 'bound-zotero-abbreviation-handler.js' });
    requireValue(typeof context.Zotero.Cite.getAbbreviation === 'function', 'Bound source did not define abbreviation handler');
  }
  vm.runInContext(`
    var rendererSystem = {
      retrieveLocale: function (lang) {
        if (!Object.prototype.hasOwnProperty.call(input.locales, lang)) throw new Error('Missing bound locale: ' + lang);
        return input.locales[lang];
      },
      retrieveItem: function (id) {
        if (!Object.prototype.hasOwnProperty.call(input.itemsByID, String(id))) throw new Error('Missing bound item: ' + id);
        return JSON.parse(JSON.stringify(input.itemsByID[String(id)]));
      }
    };
    if (input.automatic_journal_abbreviations) {
      rendererSystem.getAbbreviation = function (listname, obj, jurisdiction, category, key) {
        if (typeof category !== 'string') {
          throw new Error('Malformed abbreviation request');
        }
        if (!abbreviationCategories.has(category) && !abbreviationCategories.has(category + '-word')) {
          // Zotero's bound handler explicitly does nothing for such categories
          // (including ordinary title). Preserve that defined behavior.
          abbreviationCategoriesWithoutData.add(category);
          return Zotero.Cite.getAbbreviation(listname, obj, jurisdiction, category, key);
        }
        if (typeof jurisdiction !== 'string' || !['string', 'number'].includes(typeof key)) throw new Error('Malformed handled abbreviation request');
        return Zotero.Cite.getAbbreviation(listname, obj, jurisdiction, category, key);
      };
    }
    var processor = input.locale
      ? new CSL.Engine(rendererSystem, input.style_xml, input.locale)
      : new CSL.Engine(rendererSystem, input.style_xml);
    processor.setOutputFormat('html');
    var finalCitations = Array(input.citations.length).fill(null);
    for (var index = 0; index < input.citations.length; index++) {
      var citation = JSON.parse(JSON.stringify(input.citations[index]));
      var updates = processor.appendCitationCluster(citation);
      if (!Array.isArray(updates)) throw new Error('Invalid CSL update list');
      for (var update of updates) {
        var position = update[0], rendered = update[1], citationID = update[2];
        if (!Number.isInteger(position) || position < 0 || position > index
          || citationID !== input.citations[position].citationID || typeof rendered !== 'string') {
          throw new Error('Unbound CSL citation update');
        }
        finalCitations[position] = { citationID: citationID, html: rendered };
      }
    }
    if (finalCitations.some(value => value === null)) throw new Error('Missing final CSL citation output');
    var bibliography = processor.makeBibliography();
    if (!bibliography || !Array.isArray(bibliography[1]) || !bibliography[0]
      || !Array.isArray(bibliography[0].entry_ids)
      || bibliography[0].entry_ids.length !== bibliography[1].length) {
      throw new Error('Selected style did not produce a bound bibliography');
    }
    var renderResult = {
      citations: finalCitations,
      bibliography: { entry_ids: bibliography[0].entry_ids, entriesHTML: bibliography[1],
        layout: { line_spacing: bibliography[0].linespacing, entry_spacing: bibliography[0].entryspacing,
          second_field_align: bibliography[0]['second-field-align'], hanging_indent: bibliography[0].hangingindent } },
      processor_version: String(CSL.PROCESSOR_VERSION),
      warnings: rendererWarnings,
      abbreviation_categories_without_data: Array.from(abbreviationCategoriesWithoutData).sort(),
      scope: 'Expected CSL HTML only; no Word/Zotero actions or fields created.'
    };
  `, context, { timeout: 30000, filename: 'read-only-csl-render.js' });
  return JSON.parse(JSON.stringify(context.renderResult));
}

function main() {
  try {
    const raw = fs.readFileSync(0, 'utf8');
    requireValue(Buffer.byteLength(raw) <= 32 * 1024 * 1024, 'Expected-render input exceeds bounded size');
    const result = render(JSON.parse(raw));
    process.stdout.write(JSON.stringify(result) + '\n');
  } catch (error) {
    process.stdout.write(JSON.stringify({ result: 'HOLD', error: String(error && error.message || error) }) + '\n');
    process.exitCode = 2;
  }
}

if (require.main === module) main();
module.exports = { render };
