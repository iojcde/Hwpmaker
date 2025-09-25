// Main entry point for the Hwpmaker library
import { generateSectionXml } from './generators.js';
import { escapeXml, createIdGenerator } from './utils.js';
import { buildLinesegArray } from './builders/paragraphs.js';
import { DEFAULT_CHOICE_NUMERALS, DEFAULT_MINIFY_OPTIONS } from './constants.js';

export {
  generateSectionXml,
  escapeXml,
  createIdGenerator,
  buildLinesegArray,
  DEFAULT_CHOICE_NUMERALS,
  DEFAULT_MINIFY_OPTIONS
};