/**
 * @fileoverview TemplateEngine - Text placeholder replacement engine.
 *
 * Handles {{placeholder}} replacement in slide XML.
 *
 * Challenge with OpenXML text replacement:
 * ─────────────────────────────────────────────────────────────────
 * A simple string "{{title}}" in PowerPoint might be split across
 * MULTIPLE text run elements in the XML:
 *
 *   <a:r><a:t>{{ti</a:t></a:r>
 *   <a:r><a:t>tle}}</a:t></a:r>
 *
 * This happens because PowerPoint splits text runs when:
 *  - Spell-check marks parts of the text
 *  - Font formatting changes mid-word
 *  - The text was typed in multiple sessions
 *
 * Solution Strategy:
 * 1. Extract all <a:r><a:t>...</a:t></a:r> run sequences within each <a:p>
 * 2. Concatenate run text to find the full combined text
 * 3. Check if the combined text contains any placeholder
 * 4. If yes: merge all runs into a single run with replaced text,
 *    preserving the formatting of the FIRST run as the template
 *
 * This "text normalization" approach correctly handles fragmented placeholders.
 */

import { createLogger } from '../utils/logger.js';

const logger = createLogger('TemplateEngine');

/**
 * Default placeholder pattern: {{key}}
 * Can be overridden per-call.
 */
const DEFAULT_PLACEHOLDER_PATTERN = /\{\{([^{}]+)\}\}/g;

/**
 * @class TemplateEngine
 * @description Handles text placeholder replacement in OpenXML slide XML.
 *
 * Implements the text-normalization strategy to handle fragmented placeholders.
 */
export class TemplateEngine {
  /** @private @type {XMLParser} */
  #xmlParser;

  /**
   * @param {XMLParser} xmlParser
   */
  constructor(xmlParser) {
    this.#xmlParser = xmlParser;
  }

  /**
   * Replaces all placeholders in a slide XML string.
   * Uses the text-normalization strategy to handle fragmented text runs.
   *
   * @param {string} slideXml - Raw slide XML.
   * @param {Object.<string, string>} replacements - Placeholder → value map.
   * @param {RegExp} [pattern] - Custom placeholder pattern. Defaults to {{key}}.
   * @returns {string} Modified slide XML with placeholders replaced.
   *
   * @example
   * const updated = engine.replaceTextInXml(slideXml, {
   *   '{{name}}': 'John Doe',
   *   '{{date}}': '2026-01-01'
   * });
   */
  replaceTextInXml(slideXml, replacements, pattern = DEFAULT_PLACEHOLDER_PATTERN) {
    if (!replacements || Object.keys(replacements).length === 0) {
      return slideXml;
    }

    logger.debug(`Replacing ${Object.keys(replacements).length} placeholder(s)`);

    // Step 1: Process paragraph by paragraph to handle fragmented runs
    let updated = this.#processParagraphs(slideXml, replacements);

    // Step 2: Simple direct replacement for any remaining unfragmented placeholders
    for (const [placeholder, value] of Object.entries(replacements)) {
      const escaped = this.#escapeXml(String(value));
      const placeholderEscaped = this.#escapeXml(placeholder);

      // Replace the XML-escaped form (e.g., {{name}} as {{name}})
      updated = updated.split(placeholderEscaped).join(escaped);
      // Replace the plain form (in case it's not escaped in the XML)
      updated = updated.split(placeholder).join(escaped);
    }

    return updated;
  }

  /**
   * Processes all paragraphs (<a:p>) in the slide XML, normalizing text runs
   * within each paragraph to fix fragmented placeholders.
   *
   * @private
   * @param {string} slideXml
   * @param {Object.<string, string>} replacements
   * @returns {string}
   */
  #processParagraphs(slideXml, replacements) {
    // Find all <a:p>...</a:p> paragraphs
    let updated = slideXml;
    let offset = 0;

    const paragraphPattern = /<a:p>([\s\S]*?)<\/a:p>/g;
    let match;

    while ((match = paragraphPattern.exec(slideXml)) !== null) {
      const paragraphXml = match[0];
      const processedParagraph = this.#processParagraph(paragraphXml, replacements);

      if (processedParagraph !== paragraphXml) {
        // Replace in the updated string (adjust for offset changes)
        const start = updated.indexOf(paragraphXml, match.index + offset - (match.index));

        // More reliable: replace from the beginning of the current search area
        updated = updated.substring(0, match.index + offset) +
          processedParagraph +
          updated.substring(match.index + offset + paragraphXml.length);

        offset += processedParagraph.length - paragraphXml.length;
      }
    }

    return updated;
  }

  /**
   * Processes a single paragraph, normalizing runs and replacing placeholders.
   *
   * @private
   * @param {string} paragraphXml - XML of a single <a:p> element.
   * @param {Object.<string, string>} replacements
   * @returns {string} Updated paragraph XML.
   */
  #processParagraph(paragraphXml, replacements) {
    // Extract all text runs from this paragraph
    const runs = this.#extractRuns(paragraphXml);

    if (runs.length === 0) return paragraphXml;

    // Combine text from all runs
    const combinedText = runs.map(r => r.text).join('');

    // Check if any placeholder appears in the combined text
    let hasPlaceholder = false;
    for (const placeholder of Object.keys(replacements)) {
      if (combinedText.includes(placeholder)) {
        hasPlaceholder = true;
        break;
      }
    }

    if (!hasPlaceholder) return paragraphXml;

    // Perform replacement on combined text
    let replacedText = combinedText;
    for (const [placeholder, value] of Object.entries(replacements)) {
      replacedText = replacedText.split(placeholder).join(String(value));
    }

    // Rebuild the paragraph: merge all runs into a single run using first run's format
    return this.#mergeRunsWithText(paragraphXml, runs, replacedText);
  }

  /**
   * Extracts all text runs from a paragraph's XML.
   * Returns run XML and the extracted text content.
   *
   * @private
   * @param {string} paragraphXml
   * @returns {Array<{xml: string, text: string, start: number, end: number}>}
   */
  #extractRuns(paragraphXml) {
    const runs = [];
    const runPattern = /(<a:r(?:\s[^>]*)?>)([\s\S]*?)(<\/a:r>)/g;
    let match;

    while ((match = runPattern.exec(paragraphXml)) !== null) {
      const runXml = match[0];
      // Extract text from <a:t>...</a:t> within this run
      const tMatch = /<a:t>([\s\S]*?)<\/a:t>/.exec(runXml);
      const text = tMatch ? this.#unescapeXml(tMatch[1]) : '';

      runs.push({
        xml: runXml,
        text,
        start: match.index,
        end: match.index + runXml.length,
      });
    }

    return runs;
  }

  /**
   * Merges all runs in a paragraph into a single run using the first run's
   * formatting as the template, with the replaced text as content.
   *
   * @private
   * @param {string} paragraphXml
   * @param {Array} runs - Extracted run info.
   * @param {string} newText - New text to inject.
   * @returns {string} Updated paragraph XML.
   */
  #mergeRunsWithText(paragraphXml, runs, newText) {
    if (runs.length === 0) return paragraphXml;

    // Use the first run as the format template
    const firstRun = runs[0];

    // Build the replacement run: first run's format + new text
    const mergedRunXml = this.#setRunText(firstRun.xml, this.#escapeXml(newText));

    // Build new paragraph:
    // Keep everything before first run, insert merged run, remove the rest,
    // keep everything after last run
    const firstRunStart = firstRun.start;
    const lastRunEnd = runs[runs.length - 1].end;

    return (
      paragraphXml.substring(0, firstRunStart) +
      mergedRunXml +
      paragraphXml.substring(lastRunEnd)
    );
  }

  /**
   * Replaces the text content of a text run XML.
   *
   * @private
   * @param {string} runXml - Run XML string.
   * @param {string} text - New text content (already XML-escaped).
   * @returns {string} Updated run XML.
   */
  #setRunText(runXml, text) {
    const tPattern = /(<a:t>)([\s\S]*?)(<\/a:t>)/;
    if (tPattern.test(runXml)) {
      return runXml.replace(tPattern, `$1${text}$3`);
    }
    // If no <a:t>, add one before </a:r>
    return runXml.replace('</a:r>', `<a:t>${text}</a:t></a:r>`);
  }

  /**
   * Checks if a string contains any placeholder from the map.
   *
   * @param {string} text
   * @param {Object.<string, string>} replacements
   * @returns {boolean}
   */
  containsPlaceholders(text, replacements) {
    return Object.keys(replacements).some(p => text.includes(p));
  }

  /**
   * Extracts all unique placeholder keys from an XML string.
   *
   * @param {string} xml - Slide XML.
   * @param {RegExp} [pattern] - Placeholder pattern.
   * @returns {string[]} Array of placeholder keys found.
   *
   * @example
   * engine.extractPlaceholders(slideXml);
   * // → ['{{title}}', '{{date}}', '{{company}}']
   */
  extractPlaceholders(xml, pattern = DEFAULT_PLACEHOLDER_PATTERN) {
    const placeholders = new Set();
    const textPattern = /<a:t>([\s\S]*?)<\/a:t>/g;
    let match;

    // Extract text content first, then find placeholders
    const allText = [];
    while ((match = textPattern.exec(xml)) !== null) {
      allText.push(match[1]);
    }

    const combined = allText.join('');
    const plPattern = new RegExp(pattern.source, 'g');
    let plMatch;
    while ((plMatch = plPattern.exec(combined)) !== null) {
      placeholders.add(plMatch[0]);
    }

    return Array.from(placeholders);
  }

  /**
   * Escapes XML special characters.
   * @private
   * @param {string} str
   * @returns {string}
   */
  #escapeXml(str) {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  /**
   * Unescapes XML entities.
   * @private
   * @param {string} str
   * @returns {string}
   */
  #unescapeXml(str) {
    return str
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'");
  }
}
