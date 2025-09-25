import { buildParagraph, buildLinesegArray } from './paragraphs.js';
import { buildContextDataTable } from './tables.js';

function normalizeContextTableCell(value) {
  if (value === null || value === undefined) {
    return '';
  }

  if (typeof value === 'object' && value.label !== undefined) {
    return String(value.label);
  }

  return String(value);
}

function normalizeContextTableHeader(header, index) {
  if (header && typeof header === 'object') {
    const label = header.label ?? header.text ?? header.name ?? '';
    const key = header.key ?? header.id ?? (label ? null : `col${index}`);
    return { label: String(label), key };
  }

  const label = String(header ?? '');
  return { label, key: label ? null : `col${index}` };
}

function normalizeContextEntry(entry) {
  if (entry && typeof entry === 'object') {
    if (entry.type === 'table' || entry.table || Array.isArray(entry.rows) || Array.isArray(entry.headers)) {
      const tableSource = entry.table && typeof entry.table === 'object' ? entry.table : entry;
      const headerDefinitions = Array.isArray(tableSource.headers)
        ? tableSource.headers.map((header, index) => normalizeContextTableHeader(header, index))
        : null;

      const headerLabels = headerDefinitions ? headerDefinitions.map((header) => header.label) : null;
      const headerKeys = headerDefinitions ? headerDefinitions.map((header) => header.key) : null;

      const rawRows = Array.isArray(tableSource.rows) ? tableSource.rows : [];
      const normalizedRowArrays = rawRows.map((row) => {
        if (Array.isArray(row)) {
          return row.map(normalizeContextTableCell);
        }

        if (row && typeof row === 'object') {
          if (headerKeys && headerKeys.some((key) => key !== null)) {
            return headerKeys.map((key, idx) => {
              if (key) {
                return normalizeContextTableCell(row[key]);
              }

              const headerLabel = headerLabels ? headerLabels[idx] : null;
              if (headerLabel && Object.prototype.hasOwnProperty.call(row, headerLabel)) {
                return normalizeContextTableCell(row[headerLabel]);
              }

              return '';
            });
          }

          return Object.values(row).map(normalizeContextTableCell);
        }

        return [normalizeContextTableCell(row)];
      });

      let columnCount = headerLabels && headerLabels.length > 0
        ? headerLabels.length
        : normalizedRowArrays.reduce((max, row) => Math.max(max, row.length), 0);

      if (columnCount === 0 && normalizedRowArrays.length > 0) {
        columnCount = normalizedRowArrays[0]?.length ?? 0;
      }

      const rows = normalizedRowArrays.map((row) => {
        const trimmed = row.slice(0, columnCount);
        while (trimmed.length < columnCount) {
          trimmed.push('');
        }
        return trimmed;
      });

      return {
        type: 'table',
        label: entry.label ?? null,
        headers: headerLabels,
        rows,
        columnCount
      };
    }

    return {
      type: 'text',
      label: entry.label ?? null,
      text: entry.text ?? entry.content ?? entry.description ?? ''
    };
  }

  return {
    type: 'text',
    label: null,
    text: String(entry ?? '')
  };
}

function estimateContextRowHeight(text) {
  if (!text) {
    return 1150;
  }

  const textLength = String(text).length;
  const baseHeight = 1150;
  const heightPerCharacter = 30;

  const estimatedHeight = baseHeight + Math.floor(textLength / 30) * heightPerCharacter;
  return Math.max(baseHeight, Math.min(estimatedHeight, 5000));
}

function buildContextTable({
  paragraphId,
  tableId,
  entries,
  paragraphIdFactory
}) {
  if (!entries || entries.length === 0) {
    return '';
  }
  const rowHeights = [];

  const rows = entries.map((entry, rowIndex) => {
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
    `      <hp:tbl id="${tableId}" zOrder="12" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="1" rowCnt="${entries.length}" colCnt="1" cellSpacing="0" borderFillIDRef="7" noAdjust="0">`,
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

function buildUnifiedContextTable({
  paragraphId,
  tableId,
  entries,
  paragraphIdFactory,
  tableIdFactory
}) {
  if (!entries || entries.length === 0) {
    return '';
  }

  const rowHeights = [];
  const rows = [];

  entries.forEach((entry, rowIndex) => {
    if (entry.type === 'table') {
      // Handle data table entries - embed them as a nested table within a cell
      const labelText = entry.label || '';
      const rowHeight = estimateContextRowHeight(labelText) + 3000; // Extra height for nested table
      rowHeights.push(rowHeight);

      const nestedTableId = tableIdFactory();
      const nestedTable = buildContextDataTable({
        entry,
        paragraphId: paragraphIdFactory(),
        tableId: nestedTableId,
        paragraphIdFactory
      });

      let cellContent;
      if (labelText) {
        const labelParagraphId = paragraphIdFactory();
        const labelParagraph = buildParagraph({
          id: labelParagraphId,
          paraPrIDRef: '60',
          styleIDRef: '44',
          charPrIDRef: '46',
          text: labelText,
          lineSegOptions: {
            spacing: 516,
            horzsize: 28908,
            flags: '1441792'
          },
          indent: '              '
        });
        cellContent = [labelParagraph, nestedTable].join('\n');
      } else {
        cellContent = nestedTable;
      }

      rows.push([
        '        <hp:tr>',
        '          <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="6">',
        '            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">',
        cellContent,
        '            </hp:subList>',
        `            <hp:cellAddr colAddr="0" rowAddr="${rowIndex}" />`,
        '            <hp:cellSpan colSpan="1" rowSpan="1" />',
        `            <hp:cellSz width="30611" height="${rowHeight}" />`,
        '            <hp:cellMargin left="850" right="850" top="850" bottom="850" />',
        '          </hp:tc>',
        '        </hp:tr>'
      ].join('\n'));
    } else {
      // Handle text entries - same as before
      const entryParagraphId = paragraphIdFactory();
      const displayText = entry.label ? `${entry.label} : ${entry.text}` : entry.text;
      const rowHeight = estimateContextRowHeight(displayText) + 1700;
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

      rows.push([
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
      ].join('\n'));
    }
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
    `      <hp:tbl id="${tableId}" zOrder="12" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="1" rowCnt="${entries.length}" colCnt="1" cellSpacing="0" borderFillIDRef="7" noAdjust="0">`,
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

function buildContextBlocks({
  entries,
  paragraphIdFactory,
  tableIdFactory
}) {
  if (!entries || entries.length === 0) {
    return [];
  }

  const normalizedEntries = entries.map(normalizeContextEntry);
  
  // Create a single context table that contains all entries (text and data tables)
  const paragraphId = paragraphIdFactory();
  const tableId = tableIdFactory();
  const contextTable = buildUnifiedContextTable({
    paragraphId,
    tableId,
    entries: normalizedEntries,
    paragraphIdFactory,
    tableIdFactory
  });

  return contextTable ? [contextTable] : [];
}

export {
  normalizeContextTableCell,
  normalizeContextTableHeader, 
  normalizeContextEntry,
  estimateContextRowHeight,
  buildContextTable,
  buildUnifiedContextTable,
  buildContextBlocks
};