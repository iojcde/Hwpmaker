const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const SECTION_NAMESPACES = [
  'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app"',
  'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"',
  'xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph"',
  'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"',
  'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core"',
  'xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"',
  'xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history"',
  'xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page"',
  'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"',
  'xmlns:dc="http://purl.org/dc/elements/1.1/"',
  'xmlns:opf="http://www.idpf.org/2007/opf/"',
  'xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart"',
  'xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar"',
  'xmlns:epub="http://www.idpf.org/2007/ops"',
  'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'
].join(' ');

const SECTION_OPEN = `${XML_DECLARATION}\n<hs:sec ${SECTION_NAMESPACES}>`;
const SECTION_CLOSE = '</hs:sec>';

const DEFAULT_CHOICE_NUMERALS = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩'];
const DEFAULT_STATEMENT_TITLE = '<보 기>';
const DEFAULT_MINIFY_OPTIONS = {
  removeComments: true,
  removeWhitespaceBetweenTags: true,
  considerPreserveWhitespace: true,
  collapseWhitespaceInTags: true,
  collapseEmptyElements: true,
  trimWhitespaceFromTexts: false,
  collapseWhitespaceInTexts: false,
  collapseWhitespaceInProlog: true,
  collapseWhitespaceInDocType: true,
  removeSchemaLocationAttributes: false,
  removeUnnecessaryStandaloneDeclaration: true,
  removeUnusedNamespaces: false,
  removeUnusedDefaultNamespace: false,
  shortenNamespaces: false,
  ignoreCData: true
};
const CHOICE_LAYOUTS = {
  PARAGRAPH: 'paragraph',
  TABLE: 'table'
};

export {
  XML_DECLARATION,
  SECTION_NAMESPACES,
  SECTION_OPEN,
  SECTION_CLOSE,
  DEFAULT_CHOICE_NUMERALS,
  DEFAULT_STATEMENT_TITLE,
  DEFAULT_MINIFY_OPTIONS,
  CHOICE_LAYOUTS
};