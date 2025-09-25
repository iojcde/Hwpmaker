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
const CHOICE_LAYOUTS = {
  PARAGRAPH: 'paragraph',
  TABLE: 'table'
};

function escapeXml(value) {
  if (value === undefined || value === null) {
    return '';
  }
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function createIdGenerator(start = 2147483648) {
  let current = start;
  return () => {
    current += 1;
    return String(current);
  };
}

function buildLinesegArray({
  textpos = 0,
  vertpos = 0,
  vertsize = 1150,
  textheight = 1150,
  baseline = 978,
  spacing = 460,
  horzpos = 0,
  horzsize = 31688,
  flags = '393216'
} = {}, indent = '    ') {
  const innerIndent = `${indent}  `;
  return `${indent}<hp:linesegarray>\n${innerIndent}<hp:lineseg textpos="${textpos}" vertpos="${vertpos}" vertsize="${vertsize}" textheight="${textheight}" baseline="${baseline}" spacing="${spacing}" horzpos="${horzpos}" horzsize="${horzsize}" flags="${flags}" />\n${indent}</hp:linesegarray>`;
}

function buildParagraph({
  id,
  paraPrIDRef,
  styleIDRef,
  charPrIDRef,
  text,
  lineSegOptions,
  includeEmptyRun = false,
  indent = '  '
}) {
  const escaped = escapeXml(text);
  const runIndent = `${indent}  `;
  const textIndent = `${runIndent}  `;
  const lineseg = buildLinesegArray(lineSegOptions, `${indent}  `);

  let runText = '';
  if (escaped.length > 0) {
    runText = `${textIndent}<hp:t>${escaped}</hp:t>\n`;
  } else if (includeEmptyRun) {
    runText = `${textIndent}<hp:t />\n`;
  }

  const runBlock = `${runIndent}<hp:run charPrIDRef="${charPrIDRef}">\n${runText}${runIndent}</hp:run>`;

  return `${indent}<hp:p id="${id}" paraPrIDRef="${paraPrIDRef}" styleIDRef="${styleIDRef}" pageBreak="0" columnBreak="0" merged="0">\n${runBlock}\n${lineseg}\n${indent}</hp:p>`;
}

function buildSectionOpeningParagraph({ prompt, paragraphId, tableId }) {
  const escapedPrompt = escapeXml(prompt);
  return `  <hp:p id="${paragraphId}" paraPrIDRef="58" styleIDRef="1" pageBreak="0" columnBreak="0" merged="0">\n    <hp:run charPrIDRef="3">\n      <hp:ctrl>\n        <hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="2" sameSz="1" sameGap="3120" />\n      </hp:ctrl>\n      <hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="1" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="4">\n        <hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0" />\n        <hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0" />\n        <hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0" />\n        <hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0" />\n        <hp:pagePr landscape="WIDELY" width="77102" height="111685" gutterType="LEFT_RIGHT">\n          <hp:margin header="4960" footer="3401" gutter="0" left="5300" right="5300" top="6236" bottom="5952" />\n        </hp:pagePr>\n        <hp:footNotePr>\n          <hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0" />\n          <hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000" />\n          <hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850" />\n          <hp:numbering type="CONTINUOUS" newNum="1" />\n          <hp:placement place="EACH_COLUMN" beneathText="0" />\n        </hp:footNotePr>\n        <hp:endNotePr>\n          <hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0" />\n          <hp:noteLine length="14692344" type="SOLID" width="0.12 mm" color="#000000" />\n          <hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850" />\n          <hp:numbering type="CONTINUOUS" newNum="1" />\n          <hp:placement place="END_OF_DOCUMENT" beneathText="0" />\n        </hp:endNotePr>\n        <hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">\n          <hp:offset left="1417" right="1417" top="1417" bottom="1417" />\n        </hp:pageBorderFill>\n        <hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">\n          <hp:offset left="1417" right="1417" top="1417" bottom="1417" />\n        </hp:pageBorderFill>\n        <hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">\n          <hp:offset left="1417" right="1417" top="1417" bottom="1417" />\n        </hp:pageBorderFill>\n        <hp:masterPage idRef="masterpage0" />\n        <hp:masterPage idRef="masterpage1" />\n        <hp:masterPage idRef="masterpage2" />\n        <hp:masterPage idRef="masterpage3" />\n      </hp:secPr>\n    </hp:run>\n    <hp:run charPrIDRef="3">\n      <hp:ctrl>\n        <hp:newNum num="1" numType="PAGE" />\n      </hp:ctrl>\n      <hp:ctrl>\n        <hp:footer id="3" applyPageType="BOTH">\n          <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="BOTTOM" linkListIDRef="0" linkListNextIDRef="0" textWidth="66502" textHeight="3401" hasTextRef="0" hasNumRef="0">\n            <hp:p id="2147483648" paraPrIDRef="56" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">\n              <hp:run charPrIDRef="54" />\n              <hp:linesegarray>\n                <hp:lineseg textpos="0" vertpos="0" vertsize="1150" textheight="1150" baseline="1150" spacing="460" horzpos="0" horzsize="66500" flags="393216" />\n              </hp:linesegarray>\n            </hp:p>\n          </hp:subList>\n        </hp:footer>\n      </hp:ctrl>\n      <hp:ctrl>\n        <hp:header id="1" applyPageType="ODD">\n          <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="66502" textHeight="4960" hasTextRef="0" hasNumRef="0">\n            <hp:p id="0" paraPrIDRef="45" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">\n              <hp:run charPrIDRef="54" />\n              <hp:linesegarray>\n                <hp:lineseg textpos="0" vertpos="0" vertsize="1150" textheight="1150" baseline="978" spacing="460" horzpos="0" horzsize="66500" flags="393216" />\n              </hp:linesegarray>\n            </hp:p>\n          </hp:subList>\n        </hp:header>\n      </hp:ctrl>\n      <hp:ctrl>\n        <hp:header id="2" applyPageType="EVEN">\n          <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="66502" textHeight="4960" hasTextRef="0" hasNumRef="0">\n            <hp:p id="0" paraPrIDRef="28" styleIDRef="33" pageBreak="0" columnBreak="0" merged="0">\n              <hp:run charPrIDRef="25" />\n              <hp:linesegarray>\n                <hp:lineseg textpos="0" vertpos="0" vertsize="900" textheight="900" baseline="765" spacing="452" horzpos="0" horzsize="66500" flags="393216" />\n              </hp:linesegarray>\n            </hp:p>\n          </hp:subList>\n        </hp:header>\n      </hp:ctrl>\n      <hp:tbl id="${tableId}" zOrder="8" numberingType="TABLE" textWrap="SQUARE" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="1" rowCnt="1" colCnt="1" cellSpacing="0" borderFillIDRef="3" noAdjust="0">\n        <hp:sz width="66472" widthRelTo="ABSOLUTE" height="13888" heightRelTo="ABSOLUTE" protect="0" />\n        <hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PAPER" horzRelTo="PAGE" vertAlign="TOP" horzAlign="CENTER" vertOffset="5215" horzOffset="0" />\n        <hp:outMargin left="0" right="0" top="0" bottom="1134" />\n        <hp:inMargin left="0" right="0" top="0" bottom="0" />\n        <hp:tr>\n          <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="20">\n            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">\n              <hp:p id="2147483648" paraPrIDRef="47" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">\n                <hp:run charPrIDRef="53" />\n                <hp:linesegarray>\n                  <hp:lineseg textpos="0" vertpos="0" vertsize="1150" textheight="1150" baseline="978" spacing="804" horzpos="0" horzsize="66472" flags="393216" />\n                </hp:linesegarray>\n              </hp:p>\n            </hp:subList>\n            <hp:cellAddr colAddr="0" rowAddr="0" />\n            <hp:cellSpan colSpan="1" rowSpan="1" />\n            <hp:cellSz width="66472" height="13888" />\n            <hp:cellMargin left="141" right="141" top="141" bottom="141" />\n          </hp:tc>\n        </hp:tr>\n      </hp:tbl>\n      <hp:ctrl>\n        <hp:pageHiding hideHeader="1" hideFooter="0" hideMasterPage="0" hideBorder="0" hideFill="0" hidePageNum="0" />\n      </hp:ctrl>\n      <hp:t>${escapedPrompt}</hp:t>\n    </hp:run>\n    <hp:linesegarray>\n      <hp:lineseg textpos="0" vertpos="9041" vertsize="1400" textheight="1400" baseline="1190" spacing="560" horzpos="0" horzsize="31688" flags="2490368" />\n      <hp:lineseg textpos="98" vertpos="11001" vertsize="1150" textheight="1150" baseline="978" spacing="460" horzpos="0" horzsize="31688" flags="1441792" />\n    </hp:linesegarray>\n  </hp:p>`;
}

function buildPromptParagraph({ prompt, paragraphId }) {
  return buildParagraph({
    id: paragraphId,
    paraPrIDRef: '58',
    styleIDRef: '1',
    charPrIDRef: '3',
    text: prompt,
    lineSegOptions: {
      vertsize: 1400,
      textheight: 1400,
      baseline: 1190,
      spacing: 560,
      horzsize: 31688,
      flags: '2490368'
    }
  });
}

function buildContextTable({
  paragraphId,
  tableId,
  entries,
  paragraphIdFactory
}) {
  const normalizedEntries = entries.map(normalizeContextEntry);
  const rowHeights = [];

  const rows = normalizedEntries.map((entry, rowIndex) => {
    const entryParagraphId = paragraphIdFactory();
    const displayText = entry.label ? `${entry.label} : ${entry.text}` : entry.text;
    const rowHeight = estimateContextRowHeight(displayText) + 1700; // include cell padding allowance
    rowHeights.push(rowHeight);

    const paragraph = buildParagraph({
      id: entryParagraphId,
      paraPrIDRef: '60',
      styleIDRef: '44',
      charPrIDRef: '46',
      text: displayText,
      lineSegOptions: {
        spacing: 516,
        horzsize: 28908,
        flags: '1441792'
      },
      indent: '              '
    });

    return [
      '        <hp:tr>',
      '          <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="6">',
      '            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">',
      paragraph,
      '            </hp:subList>',
      `            <hp:cellAddr colAddr="0" rowAddr="${rowIndex}" />`,
      '            <hp:cellSpan colSpan="1" rowSpan="1" />',
      `            <hp:cellSz width="30611" height="${rowHeight}" />`,
      '            <hp:cellMargin left="850" right="850" top="850" bottom="850" />',
      '          </hp:tc>',
      '        </hp:tr>'
    ].join('\n');
  });

  const tableHeight = rowHeights.reduce((total, height) => total + height, 1700);
  const outerLineseg = buildLinesegArray({
    vertsize: tableHeight,
    textheight: tableHeight,
    baseline: Math.max(978, tableHeight - 360),
    spacing: 460,
    horzpos: 1130,
    horzsize: 30558,
    flags: '393216'
  }, '    ');

  return [
    `  <hp:p id="${paragraphId}" paraPrIDRef="3" styleIDRef="4" pageBreak="0" columnBreak="0" merged="0">`,
    '    <hp:run charPrIDRef="61">',
    `      <hp:tbl id="${tableId}" zOrder="12" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="1" rowCnt="${normalizedEntries.length}" colCnt="1" cellSpacing="0" borderFillIDRef="7" noAdjust="0">`,
    '        <hp:sz width="30611" widthRelTo="ABSOLUTE" height="' + tableHeight + '" heightRelTo="ABSOLUTE" protect="0" />',
    '        <hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0" />',
    '        <hp:outMargin left="0" right="0" top="0" bottom="0" />',
    '        <hp:inMargin left="850" right="850" top="850" bottom="850" />',
    rows.join('\n'),
    '      </hp:tbl>',
    '    </hp:run>',
    '    <hp:run charPrIDRef="1">',
    '      <hp:t />',
    '    </hp:run>',
    outerLineseg,
    '  </hp:p>'
  ].join('\n');
}

function normalizeStatementEntry(statement) {
  if (statement && typeof statement === 'object') {
    return statement.text ?? statement.content ?? statement.description ?? '';
  }

  return statement === undefined || statement === null ? '' : String(statement);
}

function normalizeStatementTitle(title) {
  if (!title) {
    return '<보기>';
  }

  return String(title)
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>');
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
      `${runIndent}<hp:run charPrIDRef="54">`,
      `${ctrlIndent}<hp:ctrl>`,
      `${ctrlIndent}  <hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0" />`,
      `${ctrlIndent}</hp:ctrl>`,
      `${runIndent}</hp:run>`
    ].join('\n'));
  }

  runs.push([
    `${runIndent}<hp:run charPrIDRef="54">`,
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
    `${indent}<hp:p id="${paragraphId}" paraPrIDRef="59" styleIDRef="9" pageBreak="0" columnBreak="0" merged="0">`,
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
    paraPrIDRef: '64',
    styleIDRef: '0',
    charPrIDRef: '60',
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
    paraPrIDRef: '64',
    styleIDRef: '0',
    charPrIDRef: '60',
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
    paraPrIDRef: '64',
    styleIDRef: '0',
    charPrIDRef: '60',
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
    paraPrIDRef: '64',
    styleIDRef: '0',
    charPrIDRef: '60',
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
    paraPrIDRef: '66',
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
    '            <hp:cellMargin left="850" right="850" top="850" bottom="850" />',
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
    baseline: Math.max(7625, tableHeight - 200),
    spacing: 460,
    horzpos: 1100,
    horzsize: 30588,
    flags: '393216'
  }, '    ');

  return [
    `  <hp:p id="${paragraphId}" paraPrIDRef="62" styleIDRef="8" pageBreak="0" columnBreak="0" merged="0">`,
    '    <hp:run charPrIDRef="63">',
    `      <hp:tbl id="${tableId}" zOrder="22" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="1" rowCnt="3" colCnt="3" cellSpacing="0" borderFillIDRef="22" noAdjust="0">`,
    `        <hp:sz width="30557" widthRelTo="ABSOLUTE" height="${tableHeight}" heightRelTo="ABSOLUTE" protect="0" />`,
    '        <hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0" />',
    '        <hp:outMargin left="0" right="0" top="0" bottom="0" />',
    '        <hp:inMargin left="0" right="0" top="0" bottom="0" />',
    tableRows.join('\n'),
    '      </hp:tbl>',
    '      <hp:t> </hp:t>',
    '    </hp:run>',
    lineseg,
    '  </hp:p>'
  ].join('\n');
}

function buildChoiceParagraph({ paragraphId, text }) {
  return buildParagraph({
    id: paragraphId,
    paraPrIDRef: '6',
    styleIDRef: '10',
    charPrIDRef: '2',
    text,
    lineSegOptions: {
      horzpos: 1130,
      horzsize: 30558
    }
  });
}

function buildSpacerParagraph(paragraphId) {
  return buildParagraph({
    id: paragraphId,
    paraPrIDRef: '6',
    styleIDRef: '10',
    charPrIDRef: '2',
    text: '',
    lineSegOptions: {
      horzpos: 1130,
      horzsize: 30558
    },
    includeEmptyRun: true
  });
}

function formatChoiceText(choice, index, options) {
  const numerals = options.choiceNumerals ?? DEFAULT_CHOICE_NUMERALS;
  const prefix = numerals[index] ?? `${index + 1}.`;
  return `${prefix} ${choice}`;
}

function buildChoiceTable({
  paragraphId,
  tableId,
  choices,
  options,
  paragraphIdFactory
}) {
  const numerals = options.choiceNumerals ?? DEFAULT_CHOICE_NUMERALS;
  const columnCount = Math.max(1, choices.length * 2);
  const tableHeight = 1431;

  const cells = [];
  let columnIndex = 0;

  choices.forEach((choice, idx) => {
    const numeralParagraphId = paragraphIdFactory();
    const numeralParagraph = buildParagraph({
      id: numeralParagraphId,
      paraPrIDRef: '45',
      styleIDRef: '0',
      charPrIDRef: '54',
      text: numerals[idx] ?? `${idx + 1}.`,
      lineSegOptions: {
        horzsize: 1440,
        baseline: 978,
        spacing: 460,
        flags: '393216'
      },
      indent: '              '
    });

    cells.push([
      '          <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="23">',
      '            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">',
      numeralParagraph,
      '            </hp:subList>',
      `            <hp:cellAddr colAddr="${columnIndex}" rowAddr="0" />`,
      '            <hp:cellSpan colSpan="1" rowSpan="1" />',
      '            <hp:cellSz width="1379" height="1431" />',
      '            <hp:cellMargin left="0" right="0" top="0" bottom="0" />',
      '          </hp:tc>'
    ].join('\n'));

    columnIndex += 1;

    const choiceParagraphId = paragraphIdFactory();
    const choiceText = normalizeChoice(choice);
    const choiceParagraph = buildParagraph({
      id: choiceParagraphId,
      paraPrIDRef: '45',
      styleIDRef: '0',
      charPrIDRef: '54',
      text: choiceText,
      lineSegOptions: {
        horzsize: 4560,
        baseline: 978,
        spacing: 460,
        flags: '393216'
      },
      indent: '              '
    });

    cells.push([
      '          <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="23">',
      '            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">',
      choiceParagraph,
      '            </hp:subList>',
      `            <hp:cellAddr colAddr="${columnIndex}" rowAddr="0" />`,
      '            <hp:cellSpan colSpan="1" rowSpan="1" />',
      '            <hp:cellSz width="4744" height="1431" />',
      '            <hp:cellMargin left="184" right="0" top="0" bottom="0" />',
      '          </hp:tc>'
    ].join('\n'));

    columnIndex += 1;
  });

  const lineseg = buildLinesegArray({
    vertsize: tableHeight,
    textheight: tableHeight,
    baseline: 1216,
    spacing: 460,
    horzpos: 1130,
    horzsize: 30558,
    flags: '393216'
  }, '    ');

  return [
    `  <hp:p id="${paragraphId}" paraPrIDRef="63" styleIDRef="10" pageBreak="0" columnBreak="0" merged="0">`,
    '    <hp:run charPrIDRef="0">',
    `      <hp:tbl id="${tableId}" zOrder="21" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="1" rowCnt="1" colCnt="${columnCount}" cellSpacing="0" borderFillIDRef="22" noAdjust="0">`,
    '        <hp:sz width="30607" widthRelTo="ABSOLUTE" height="1431" heightRelTo="ABSOLUTE" protect="1" />',
    '        <hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0" />',
    '        <hp:outMargin left="0" right="0" top="0" bottom="0" />',
    '        <hp:inMargin left="0" right="0" top="0" bottom="0" />',
    '        <hp:tr>',
    cells.join('\n'),
    '        </hp:tr>',
    '      </hp:tbl>',
    '    </hp:run>',
    '    <hp:run charPrIDRef="65">',
    '      <hp:t />',
    '    </hp:run>',
    lineseg,
    '  </hp:p>'
  ].join('\n');
}

function normalizeChoice(choice) {
  if (choice && typeof choice === 'object') {
    return choice.text ?? choice.content ?? choice.description ?? '';
  }

  return choice === undefined || choice === null ? '' : String(choice);
}

function normalizeContextEntry(entry) {
  if (entry && typeof entry === 'object') {
    return {
      label: entry.label ?? entry.title ?? entry.name ?? null,
      text: entry.text ?? entry.content ?? entry.description ?? ''
    };
  }

  return {
    label: null,
    text: entry === undefined || entry === null ? '' : String(entry)
  };
}

function estimateContextRowHeight(text) {
  if (!text) {
    return 2600;
  }

  const normalized = String(text)
    .replace(/\r\n/g, '\n')
    .split('\n')
    .reduce((count, line) => {
      const trimmed = line.trim();
      const segments = Math.max(1, Math.ceil(trimmed.length / 28));
      return count + segments;
    }, 0);

  const totalLines = Math.max(1, normalized);
  const lineHeight = 1970;
  return Math.max(2600, totalLines * lineHeight);
}

function normalizeQuestion(question, index) {
  if (!question || typeof question !== 'object') {
    throw new Error(`Question at index ${index} must be an object.`);
  }
  if (!question.prompt) {
    throw new Error(`Question at index ${index} is missing a prompt.`);
  }
  const requestedLayout = typeof question.choiceLayout === 'string'
    ? question.choiceLayout.toLowerCase()
    : question.choiceLayout;
  const resolvedLayout = requestedLayout === CHOICE_LAYOUTS.TABLE
    ? CHOICE_LAYOUTS.TABLE
    : requestedLayout === CHOICE_LAYOUTS.PARAGRAPH
      ? CHOICE_LAYOUTS.PARAGRAPH
      : null;

  const normalized = {
    prompt: question.prompt,
    contextEntries: Array.isArray(question.contextEntries) ? question.contextEntries : [],
    statementTitle: question.statementTitle ?? question.statementsTitle ?? null,
    statements: Array.isArray(question.statements) ? question.statements : [],
    choices: Array.isArray(question.choices) ? question.choices : [],
    choiceLayout: resolvedLayout ?? (question.statementTitle ? CHOICE_LAYOUTS.TABLE : CHOICE_LAYOUTS.PARAGRAPH),
    answer: question.answer ?? null,
    explanation: question.explanation ?? null
  };
  return normalized;
}

function generateSectionXml({ questions, options = {} }) {
  if (!Array.isArray(questions) || questions.length === 0) {
    throw new Error('"questions" must be a non-empty array.');
  }

  const paragraphIdFactory = createIdGenerator(options.baseParagraphId);
  const tableIdFactory = createIdGenerator(options.baseTableId ?? 1900000000);

  const blocks = [SECTION_OPEN];

  questions.map(normalizeQuestion).forEach((question, index) => {
    if (index === 0) {
      const firstParagraphId = paragraphIdFactory();
      const firstTableId = tableIdFactory();
      blocks.push(buildSectionOpeningParagraph({
        prompt: question.prompt,
        paragraphId: firstParagraphId,
        tableId: firstTableId
      }));
    } else {
      const paragraphId = paragraphIdFactory();
      blocks.push(buildPromptParagraph({ prompt: question.prompt, paragraphId }));
    }

    if (question.contextEntries.length > 0) {
      const paragraphId = paragraphIdFactory();
      const tableId = tableIdFactory();
      blocks.push(buildContextTable({
        paragraphId,
        tableId,
        entries: question.contextEntries,
        paragraphIdFactory
      }));
    }

    if (question.statementTitle || question.statements.length > 0) {
      const paragraphId = paragraphIdFactory();
      const tableId = tableIdFactory();
      const statementBlock = buildStatementTable({
        paragraphId,
        tableId,
        title: question.statementTitle,
        statements: question.statements,
        paragraphIdFactory
      });
      if (statementBlock) {
        blocks.push(statementBlock);
      }
    }

    if (question.choices.length > 0) {
      if (question.choiceLayout === CHOICE_LAYOUTS.TABLE) {
        const paragraphId = paragraphIdFactory();
        const tableId = tableIdFactory();
        blocks.push(buildChoiceTable({
          paragraphId,
          tableId,
          choices: question.choices,
          options,
          paragraphIdFactory
        }));
      } else {
        question.choices.forEach((choice, idx) => {
          const paragraphId = paragraphIdFactory();
          const choiceText = normalizeChoice(choice);
          const choiceContent = formatChoiceText(choiceText, idx, options);
          blocks.push(buildChoiceParagraph({ paragraphId, text: choiceContent }));
        });
      }
    }

    if (question.answer) {
      const paragraphId = paragraphIdFactory();
      blocks.push(buildParagraph({
        id: paragraphId,
        paraPrIDRef: '46',
        styleIDRef: '0',
        charPrIDRef: '49',
        text: `정답: ${question.answer}`,
        lineSegOptions: {
          horzsize: 31688
        }
      }));
    }

    if (question.explanation) {
      const paragraphId = paragraphIdFactory();
      blocks.push(buildParagraph({
        id: paragraphId,
        paraPrIDRef: '58',
        styleIDRef: '1',
        charPrIDRef: '3',
        text: question.explanation,
        lineSegOptions: {
          horzsize: 31688,
          flags: '1441792'
        }
      }));
    }

    const spacerCount = options.spacersPerQuestion ?? 1;
    for (let i = 0; i < spacerCount; i += 1) {
      blocks.push(buildSpacerParagraph(paragraphIdFactory()));
    }
  });

  blocks.push(SECTION_CLOSE);
  return `${blocks.join('\n')}`;
}

export {
  generateSectionXml,
  escapeXml,
  createIdGenerator,
  buildLinesegArray,
  DEFAULT_CHOICE_NUMERALS
};
