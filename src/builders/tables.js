import { buildParagraph, buildLinesegArray } from './paragraphs.js';

function estimateDataTableRowHeight(cells) {
  if (!Array.isArray(cells) || cells.length === 0) {
    return 1150;
  }

  const maxLength = Math.max(...cells.map((cell) => String(cell || '').length));
  const baseHeight = 1150;
  const heightPerCharacter = 20;
  return Math.max(baseHeight, baseHeight + Math.floor(maxLength / 20) * heightPerCharacter);
}

function buildContextDataTable({
  entry,
  paragraphId,
  tableId,
  paragraphIdFactory
}) {
  const columnCount = entry.columnCount ?? (entry.headers ? entry.headers.length : 0);
  const hasHeaders = Array.isArray(entry.headers) && entry.headers.length > 0;

  if (!columnCount || columnCount <= 0 || (!hasHeaders && entry.rows.length === 0)) {
    return '';
  }

  const totalWidth = 30611;
  const baseWidth = Math.floor(totalWidth / columnCount);
  const columnWidths = Array.from({ length: columnCount }, () => baseWidth);
  columnWidths[columnCount - 1] += totalWidth - baseWidth * columnCount;

  const rowHeights = [];
  const tableRows = [];

  if (hasHeaders) {
    const headerHeight = estimateDataTableRowHeight(entry.headers);
    rowHeights.push(headerHeight);

    const headerCells = entry.headers.map((headerText, colIndex) => {
      const cellParagraphId = paragraphIdFactory();
      const paragraph = buildParagraph({
        id: cellParagraphId,
        paraPrIDRef: '45',
        styleIDRef: '0',
        charPrIDRef: '54',
        text: headerText,
        lineSegOptions: {
          spacing: 516,
          horzsize: Math.max(1800, columnWidths[colIndex] - 1500),
          flags: '1441792'
        },
        indent: '              '
      });

      return [
        '          <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="6">',
        '            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">',
        paragraph,
        '            </hp:subList>',
        `            <hp:cellAddr colAddr="${colIndex}" rowAddr="0" />`,
        '            <hp:cellSpan colSpan="1" rowSpan="1" />',
        `            <hp:cellSz width="${columnWidths[colIndex]}" height="${headerHeight}" />`,
        '            <hp:cellMargin left="850" right="850" top="850" bottom="850" />',
        '          </hp:tc>'
      ].join('\n');
    });

    tableRows.push([
      '        <hp:tr>',
      headerCells.join('\n'),
      '        </hp:tr>'
    ].join('\n'));
  }

  entry.rows.forEach((row, rowIndex) => {
    const rowHeight = estimateDataTableRowHeight(row);
    rowHeights.push(rowHeight);
    const rowAddress = rowIndex + (hasHeaders ? 1 : 0);

    const cells = row.map((cellText, colIndex) => {
      const cellParagraphId = paragraphIdFactory();
      const paragraph = buildParagraph({
        id: cellParagraphId,
        paraPrIDRef: '60',
        styleIDRef: '44',
        charPrIDRef: '46',
        text: cellText,
        lineSegOptions: {
          spacing: 516,
          horzsize: Math.max(1800, columnWidths[colIndex] - 1500),
          flags: '1441792'
        },
        indent: '              '
      });

      return [
        '          <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="6">',
        '            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">',
        paragraph,
        '            </hp:subList>',
        `            <hp:cellAddr colAddr="${colIndex}" rowAddr="${rowAddress}" />`,
        '            <hp:cellSpan colSpan="1" rowSpan="1" />',
        `            <hp:cellSz width="${columnWidths[colIndex]}" height="${rowHeight}" />`,
        '            <hp:cellMargin left="850" right="850" top="850" bottom="850" />',
        '          </hp:tc>'
      ].join('\n');
    });

    tableRows.push([
      '        <hp:tr>',
      cells.join('\n'),
      '        </hp:tr>'
    ].join('\n'));
  });

  if (tableRows.length === 0) {
    return '';
  }

  const tableHeight = rowHeights.reduce((total, height) => total + height, 1700);
  const lineseg = buildLinesegArray({
    vertsize: tableHeight,
    textheight: tableHeight,
    baseline: Math.max(978, tableHeight - 360),
    spacing: 460,
    horzpos: 1130,
    horzsize: 30558,
    flags: '393216'
  }, '    ');

  const totalRows = tableRows.length;

  return [
    `  <hp:p id="${paragraphId}" paraPrIDRef="3" styleIDRef="4" pageBreak="0" columnBreak="0" merged="0">`,
    '    <hp:run charPrIDRef="61">',
    `      <hp:tbl id="${tableId}" zOrder="12" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="1" rowCnt="${totalRows}" colCnt="${columnCount}" cellSpacing="0" borderFillIDRef="22" noAdjust="0">`,
    `        <hp:sz width="30611" widthRelTo="ABSOLUTE" height="${tableHeight}" heightRelTo="ABSOLUTE" protect="0" />`,
    '        <hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0" />',
    '        <hp:outMargin left="0" right="0" top="0" bottom="0" />',
    '        <hp:inMargin left="850" right="850" top="850" bottom="850" />',
    tableRows.join('\n'),
    '      </hp:tbl>',
    '    </hp:run>',
    '    <hp:run charPrIDRef="1">',
    '      <hp:t />',
    '    </hp:run>',
    lineseg,
    '  </hp:p>'
  ].join('\n');
}

export {
  estimateDataTableRowHeight,
  buildContextDataTable
};