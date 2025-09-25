import { buildParagraph, buildLinesegArray } from './paragraphs.js';
import { escapeXml } from '../utils.js';
import { DEFAULT_STATEMENT_TITLE } from '../constants.js';

function normalizeStatementEntry(statement) {
  if (statement && typeof statement === 'object') {
    return statement.text ?? statement.content ?? statement.description ?? '';
  }

  return statement === undefined || statement === null ? '' : String(statement);
}

function normalizeStatementTitle(title) {
  if (title && typeof title === 'object') {
    return title.text ?? title.content ?? title.label ?? '';
  }

  if (title === null || title === undefined) {
    return DEFAULT_STATEMENT_TITLE;
  }

  let normalized = String(title)
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>');

  if (/^<\s*보기\s*>$/i.test(normalized) || /^<\s*보\s*기\s*>$/i.test(normalized)) {
    normalized = DEFAULT_STATEMENT_TITLE;
  }

  return normalized;
}

function estimateStatementCellHeight(statements) {
  if (!Array.isArray(statements) || statements.length === 0) {
    return 1150;
  }

  const totalLength = statements.reduce((total, statement) => {
    const normalized = normalizeStatementEntry(statement);
    return total + String(normalized).length;
  }, 0);

  const baseHeight = 1150;
  const heightPerCharacter = 15;
  const extraHeight = 200 * statements.length;

  return Math.max(baseHeight, baseHeight + Math.floor(totalLength / 50) * heightPerCharacter + extraHeight);
}

function buildStatementParagraphInTable({ paragraphId, text, includeColumnCtrl = false }) {
  const escapedText = escapeXml(text);
  const columnCtrl = includeColumnCtrl
    ? '      <hp:ctrl>\n        <hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="2" sameSz="1" sameGap="3120" />\n      </hp:ctrl>\n'
    : '';

  return `              <hp:p id="${paragraphId}" paraPrIDRef="27" styleIDRef="35" pageBreak="0" columnBreak="0" merged="0">
                <hp:run charPrIDRef="24">
${columnCtrl}                  <hp:t>${escapedText}</hp:t>
                </hp:run>
                <hp:linesegarray>
                  <hp:lineseg textpos="0" vertpos="0" vertsize="1150" textheight="1150" baseline="978" spacing="460" horzpos="0" horzsize="15844" flags="1441792" />
                </hp:linesegarray>
              </hp:p>`;
}

function buildStatementTable({
  paragraphId,
  tableId,
  title,
  statements,
  paragraphIdFactory
}) {
  if (!Array.isArray(statements) || statements.length === 0) {
    return '';
  }

  const normalizedStatements = statements.map(normalizeStatementEntry).filter(Boolean);
  if (normalizedStatements.length === 0) {
    return '';
  }

  const normalizedTitle = normalizeStatementTitle(title);
  const cellHeight = estimateStatementCellHeight(normalizedStatements);

  const titleParagraphId = paragraphIdFactory();
  const statementsParagraphId = paragraphIdFactory();

  const titleParagraph = buildStatementParagraphInTable({
    paragraphId: titleParagraphId,
    text: normalizedTitle
  });

  const statementsText = normalizedStatements.map((statement, index) => {
    const bullet = index === 0 ? 'ㄱ.' : index === 1 ? 'ㄴ.' : index === 2 ? 'ㄷ.' : `${String.fromCharCode(12593 + index)}.`;
    return `${bullet} ${statement}`;
  }).join('\n');

  const statementsParagraph = buildStatementParagraphInTable({
    paragraphId: statementsParagraphId,
    text: statementsText,
    includeColumnCtrl: true
  });

  const linesegArray = buildLinesegArray({
    vertsize: cellHeight,
    textheight: cellHeight,
    baseline: Math.max(978, cellHeight - 360),
    spacing: 460,
    horzpos: 1130,
    horzsize: 30558,
    flags: '393216'
  }, '    ');

  return [
    `  <hp:p id="${paragraphId}" paraPrIDRef="3" styleIDRef="4" pageBreak="0" columnBreak="0" merged="0">`,
    '    <hp:run charPrIDRef="61">',
    `      <hp:tbl id="${tableId}" zOrder="12" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="1" rowCnt="3" colCnt="3" cellSpacing="0" borderFillIDRef="22" noAdjust="0">`,
    `        <hp:sz width="30611" widthRelTo="ABSOLUTE" height="${cellHeight}" heightRelTo="ABSOLUTE" protect="0" />`,
    '        <hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0" />',
    '        <hp:outMargin left="0" right="0" top="0" bottom="0" />',
    '        <hp:inMargin left="850" right="850" top="850" bottom="850" />',
    '        <hp:tr>',
    '          <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="6">',
    '            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">',
    titleParagraph,
    '            </hp:subList>',
    '            <hp:cellAddr colAddr="0" rowAddr="0" />',
    '            <hp:cellSpan colSpan="1" rowSpan="3" />',
    '            <hp:cellSz width="2040" height="1150" />',
    '            <hp:cellMargin left="850" right="850" top="850" bottom="850" />',
    '          </hp:tc>',
    '          <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="6">',
    '            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">',
    statementsParagraph,
    '            </hp:subList>',
    '            <hp:cellAddr colAddr="1" rowAddr="0" />',
    '            <hp:cellSpan colSpan="2" rowSpan="3" />',
    '            <hp:cellSz width="28571" height="1150" />',
    '            <hp:cellMargin left="850" right="850" top="850" bottom="850" />',
    '          </hp:tc>',
    '        </hp:tr>',
    '        <hp:tr />',
    '        <hp:tr />',
    '      </hp:tbl>',
    '    </hp:run>',
    '    <hp:run charPrIDRef="1">',
    '      <hp:t />',
    '    </hp:run>',
    linesegArray,
    '  </hp:p>'
  ].join('\n');
}

export {
  normalizeStatementEntry,
  normalizeStatementTitle,
  estimateStatementCellHeight,
  buildStatementParagraphInTable,
  buildStatementTable
};