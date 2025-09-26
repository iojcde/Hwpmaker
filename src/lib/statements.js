import { DEFAULT_STATEMENT_TITLE } from './constants.js';
import { escapeXml } from './utils.js';
import { buildLinesegArray, buildParagraph } from './paragraph.js';

function normalizeStatementEntry(statement) {
  if (statement && typeof statement === 'object') {
    return statement.text ?? statement.content ?? statement.description ?? '';
  }

  return statement === undefined || statement === null ? '' : String(statement);
}

function normalizeStatementTitle(title) {
  if (!title) {
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
  if (statements.length === 0) {
    return 3200;
  }

  const totalLines = statements.reduce((acc, statement) => {
    const normalized = statement
      .replace(/\r\n/g, '\n')
      .split('\n')
      .reduce((count, line) => {
        const trimmed = line.trim();
        const segments = Math.max(1, Math.ceil(trimmed.length / 30));
        return count + segments;
      }, 0);

    return acc + Math.max(1, normalized);
  }, 0);

  const lineHeight = 1500;
  return Math.max(3200, totalLines * lineHeight + 400);
}

function buildStatementParagraphInTable({ paragraphId, text, includeColumnCtrl = false }) {
  const indent = '              ';
  const runIndent = `${indent}  `;
  const ctrlIndent = `${runIndent}  `;
  const runs = [];

  if (includeColumnCtrl) {
    runs.push([
      `${runIndent}<hp:run charPrIDRef="53">`,
      `${ctrlIndent}<hp:ctrl>`,
      `${ctrlIndent}  <hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0" />`,
      `${ctrlIndent}</hp:ctrl>`,
      `${runIndent}</hp:run>`
    ].join('\n'));
  }

  runs.push([
    `${runIndent}<hp:run charPrIDRef="53">`,
    `${runIndent}  <hp:t>${escapeXml(text)}</hp:t>`,
    `${runIndent}</hp:run>`
  ].join('\n'));

  const lineseg = buildLinesegArray({
    vertsize: 1150,
    textheight: 1150,
    baseline: 575,
    spacing: 460,
    horzpos: 0,
    horzsize: 28856,
    flags: '2490368'
  }, `${indent}  `);

  return [
    `${indent}<hp:p id="${paragraphId}" paraPrIDRef="63" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">`,
    runs.join('\n'),
    lineseg,
    `${indent}</hp:p>`
  ].join('\n');
}

function buildStatementTable({
  paragraphId,
  tableId,
  title,
  statements,
  paragraphIdFactory
}) {
  const normalizedStatements = statements.map(normalizeStatementEntry);
  if (normalizedStatements.length === 0 && !title) {
    return '';
  }

  const statementTitle = normalizeStatementTitle(title);

  const topLeftParagraphId = paragraphIdFactory();
  const topRightParagraphId = paragraphIdFactory();
  const secondLeftParagraphId = paragraphIdFactory();
  const secondRightParagraphId = paragraphIdFactory();
  const labelParagraphId = paragraphIdFactory();

  const blankLeftCell = buildParagraph({
    id: topLeftParagraphId,
    paraPrIDRef: '48',
    styleIDRef: '0',
    charPrIDRef: '59',
    text: '',
    lineSegOptions: {
      vertsize: 100,
      textheight: 100,
      baseline: 85,
      spacing: 40,
      horzpos: 0,
      horzsize: 13220,
      flags: '393216'
    },
    includeEmptyRun: true,
    indent: '              '
  });

  const blankRightCell = buildParagraph({
    id: topRightParagraphId,
    paraPrIDRef: '48',
    styleIDRef: '0',
    charPrIDRef: '59',
    text: '',
    lineSegOptions: {
      vertsize: 100,
      textheight: 100,
      baseline: 85,
      spacing: 40,
      horzpos: 0,
      horzsize: 12796,
      flags: '393216'
    },
    includeEmptyRun: true,
    indent: '              '
  });

  const secondRowLeftCell = buildParagraph({
    id: secondLeftParagraphId,
    paraPrIDRef: '48',
    styleIDRef: '0',
    charPrIDRef: '59',
    text: '',
    lineSegOptions: {
      vertsize: 100,
      textheight: 100,
      baseline: 85,
      spacing: 40,
      horzpos: 0,
      horzsize: 13220,
      flags: '393216'
    },
    includeEmptyRun: true,
    indent: '              '
  });

  const secondRowRightCell = buildParagraph({
    id: secondRightParagraphId,
    paraPrIDRef: '48',
    styleIDRef: '0',
    charPrIDRef: '59',
    text: '',
    lineSegOptions: {
      vertsize: 100,
      textheight: 100,
      baseline: 85,
      spacing: 40,
      horzpos: 0,
      horzsize: 12796,
      flags: '393216'
    },
    includeEmptyRun: true,
    indent: '              '
  });

  const labelParagraph = buildParagraph({
    id: labelParagraphId,
    paraPrIDRef: '61',
    styleIDRef: '38',
    charPrIDRef: '35',
    text: statementTitle,
    lineSegOptions: {
      vertsize: 1150,
      textheight: 1150,
      baseline: 978,
      spacing: 748,
      horzpos: 0,
      horzsize: 4532,
      flags: '393216'
    },
    indent: '              '
  });

  const statementParagraphs = normalizedStatements.map((statement, index) => {
    const entryParagraphId = paragraphIdFactory();
    return buildStatementParagraphInTable({
      paragraphId: entryParagraphId,
      text: statement,
      includeColumnCtrl: index === 0
    });
  });

  const statementCellHeight = estimateStatementCellHeight(normalizedStatements);
  const topRowsHeight = 1290 + 645;
  const tableHeight = topRowsHeight + statementCellHeight;

  const statementCell = [
    '        <hp:tr>',
    '          <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="12">',
    '            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">',
    statementParagraphs.join('\n'),
    '            </hp:subList>',
    '            <hp:cellAddr colAddr="0" rowAddr="2" />',
    '            <hp:cellSpan colSpan="3" rowSpan="1" />',
    `            <hp:cellSz width="30557" height="${statementCellHeight}" />`,
    '            <hp:cellMargin left="850" right="850" top="708" bottom="850" />',
    '          </hp:tc>',
    '        </hp:tr>'
  ].join('\n');

  const tableRows = [
    [
      '        <hp:tr>',
      '          <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="8">',
      '            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">',
      blankLeftCell,
      '            </hp:subList>',
      '            <hp:cellAddr colAddr="0" rowAddr="0" />',
      '            <hp:cellSpan colSpan="1" rowSpan="1" />',
      '            <hp:cellSz width="13223" height="645" />',
      '            <hp:cellMargin left="141" right="141" top="141" bottom="141" />',
      '          </hp:tc>',
      '          <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="9">',
      '            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">',
      labelParagraph,
      '            </hp:subList>',
      '            <hp:cellAddr colAddr="1" rowAddr="0" />',
      '            <hp:cellSpan colSpan="1" rowSpan="2" />',
      '            <hp:cellSz width="4535" height="1290" />',
      '            <hp:cellMargin left="141" right="141" top="141" bottom="141" />',
      '          </hp:tc>',
      '          <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="8">',
      '            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">',
      blankRightCell,
      '            </hp:subList>',
      '            <hp:cellAddr colAddr="2" rowAddr="0" />',
      '            <hp:cellSpan colSpan="1" rowSpan="1" />',
      '            <hp:cellSz width="12035" height="645" />',
      '            <hp:cellMargin left="141" right="141" top="141" bottom="141" />',
      '          </hp:tc>',
      '        </hp:tr>'
    ].join('\n'),
    [
      '        <hp:tr>',
      '          <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="10">',
      '            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">',
      secondRowLeftCell,
      '            </hp:subList>',
      '            <hp:cellAddr colAddr="0" rowAddr="1" />',
      '            <hp:cellSpan colSpan="1" rowSpan="1" />',
      '            <hp:cellSz width="13223" height="645" />',
      '            <hp:cellMargin left="141" right="141" top="141" bottom="141" />',
      '          </hp:tc>',
      '          <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="11">',
      '            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">',
      secondRowRightCell,
      '            </hp:subList>',
      '            <hp:cellAddr colAddr="2" rowAddr="1" />',
      '            <hp:cellSpan colSpan="1" rowSpan="1" />',
      '            <hp:cellSz width="12035" height="645" />',
      '            <hp:cellMargin left="141" right="141" top="141" bottom="141" />',
      '          </hp:tc>',
      '        </hp:tr>'
    ].join('\n'),
    statementCell
  ];

  const lineseg = buildLinesegArray({
    vertsize: tableHeight,
    textheight: tableHeight,
    baseline: Math.max(8872, tableHeight - 1566),
    spacing: 460,
    horzpos: 1100,
    horzsize: 30588,
    flags: '393216'
  }, '    ');

  return [
    `  <hp:p id="${paragraphId}" paraPrIDRef="59" styleIDRef="8" pageBreak="0" columnBreak="0" merged="0">`,
    '    <hp:run charPrIDRef="62">',
    `      <hp:tbl id="${tableId}" zOrder="22" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="1" rowCnt="3" colCnt="3" cellSpacing="0" borderFillIDRef="5" noAdjust="0">`,
    `        <hp:sz width="30557" widthRelTo="ABSOLUTE" height="${tableHeight}" heightRelTo="ABSOLUTE" protect="0" />`,
    '        <hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0" />',
    '        <hp:outMargin left="0" right="0" top="0" bottom="0" />',
    '        <hp:inMargin left="0" right="0" top="0" bottom="0" />',
    tableRows.join('\n'),
    '      </hp:tbl>',
    '      <hp:t />',
    '    </hp:run>',
    lineseg,
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
