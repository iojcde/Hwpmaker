import { buildParagraph, buildLinesegArray } from './paragraph.js';

const CONTEXT_TABLE_WIDTH = 30611;
const CONTEXT_OUTER_BORDER_ID = '7';
const CONTEXT_CELL_BORDER_ID = '6';
const CONTEXT_CELL_MARGIN = 850;
const NESTED_TABLE_OUT_MARGIN_TOP = 566;
const NESTED_TABLE_OUT_MARGIN_BOTTOM = 566;
const NESTED_IN_MARGIN = {
  left: 510,
  right: 510,
  top: 141,
  bottom: 141
};
const NESTED_PARAGRAPH_PARA_PR_ID = '44';
const NESTED_PARAGRAPH_STYLE_ID = '0';
const NESTED_PARAGRAPH_CHAR_PR_ID = '53';
const NESTED_TABLE_LINESEG_HORZSIZE = CONTEXT_TABLE_WIDTH - (CONTEXT_CELL_MARGIN * 2) - 3;
const CONTEXT_ROW_EXTRA_PADDING = 566;

function countVisualLines(text, { charsPerLine = 28 } = {}) {
  if (!text) {
    return 1;
  }

  const segments = String(text)
    .replace(/\r\n/g, '\n')
    .split('\n')
    .reduce((count, rawLine) => {
      const trimmed = rawLine.trim();
      if (trimmed.length === 0) {
        return count + 1;
      }

      const estimatedSegments = Math.max(1, Math.ceil(trimmed.length / charsPerLine));
      return count + estimatedSegments;
    }, 0);

  return Math.max(1, segments);
}

function estimateContextRowHeight(text) {
  const totalLines = countVisualLines(text, { charsPerLine: 32 });
  const lineHeight = 1150;
  const lineSpacing = 460;
  const basePadding = 566;

  const contentHeight = totalLines * lineHeight;
  const spacingHeight = Math.max(0, totalLines - 1) * lineSpacing;
  const computed = basePadding + contentHeight + spacingHeight;

  return Math.max(2000, computed);
}

function normalizeContextTableCell(value) {
  if (value === undefined || value === null) {
    return '';
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    return String(value);
  }

  return String(value);
}

function normalizeContextTableHeader(header, index) {
  if (header && typeof header === 'object') {
    const labelSource = header.label ?? header.text ?? header.title ?? header.name ?? index;
    const label = normalizeContextTableCell(labelSource);
    const key = header.key ?? header.field ?? header.id ?? header.name ?? null;
    return { label, key };
  }

  return {
    label: normalizeContextTableCell(header ?? index),
    key: null
  };
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
        : normalizedRowArrays.reduce((max, rowArray) => Math.max(max, rowArray.length), 0);

      if (columnCount === 0 && normalizedRowArrays.length > 0) {
        columnCount = normalizedRowArrays[0]?.length ?? 0;
      }

      const rows = normalizedRowArrays.map((rowArray) => {
        const trimmed = rowArray.slice(0, columnCount);
        while (trimmed.length < columnCount) {
          trimmed.push('');
        }
        return trimmed;
      });

      return {
        type: 'table',
        label: entry.label ?? entry.title ?? entry.name ?? null,
        text: entry.text ?? entry.description ?? entry.content ?? null,
        headers: headerLabels && headerLabels.length > 0 ? headerLabels : null,
        rows,
        columnCount
      };
    }

    return {
      type: 'text',
      label: entry.label ?? entry.title ?? entry.name ?? null,
      text: entry.text ?? entry.content ?? entry.description ?? ''
    };
  }

  return {
    type: 'text',
    label: null,
    text: entry === undefined || entry === null ? '' : String(entry)
  };
}

function formatDisplayText({ label, text } = {}) {
  const normalizedLabel = label === undefined || label === null ? '' : String(label);
  const normalizedText = text === undefined || text === null ? '' : String(text);

  const hasLabel = normalizedLabel.length > 0;
  const hasText = normalizedText.length > 0;

  if (hasLabel && hasText) {
    return `${normalizedLabel} : ${normalizedText}`;
  }

  if (hasLabel) {
    return normalizedLabel;
  }

  if (hasText) {
    return normalizedText;
  }

  return '';
}

function mergeContextEntries(entries) {
  if (!entries || entries.length === 0) {
    return [];
  }

  const merged = [];

  for (let index = 0; index < entries.length; index += 1) {
    const current = entries[index];

    if (current.type === 'text') {
      const next = entries[index + 1];
      if (next && next.type === 'table') {
        const combined = { ...next };
        const textDisplay = formatDisplayText(current);
        const tableDisplay = formatDisplayText(next);
        const displayPieces = [];

        if (textDisplay) {
          displayPieces.push(textDisplay);
        }

        if (tableDisplay && tableDisplay !== textDisplay) {
          displayPieces.push(tableDisplay);
        }

        if (displayPieces.length > 0) {
          combined.displayText = displayPieces.join('\n');
        }

        merged.push(combined);
        index += 1;
        continue;
      }
    }

    merged.push({ ...current });
  }

  return merged;
}

function estimateDataTableRowHeight(cells) {
  if (!Array.isArray(cells) || cells.length === 0) {
    return 1904;
  }

  const charsPerLine = 24;
  const lineHeight = 950;
  const lineSpacing = 380;
  const verticalPadding = 472;
  const minHeight = 1904;

  const maxCellHeight = cells.reduce((max, cell) => {
    const totalLines = countVisualLines(cell, { charsPerLine });
    const contentHeight = totalLines * lineHeight;
    const spacingHeight = Math.max(0, totalLines - 1) * lineSpacing;
    const computed = verticalPadding + contentHeight + spacingHeight;
    return Math.max(max, Math.max(minHeight, computed));
  }, minHeight);

  return maxCellHeight;
}

function buildContextDataTable({
  entry,
  paragraphIdFactory,
  tableIdFactory,
  hasFollowingEntry
}) {
  const columnCount = entry.columnCount ?? (entry.headers ? entry.headers.length : 0);
  const hasHeaders = Array.isArray(entry.headers) && entry.headers.length > 0;

  if (!columnCount || columnCount <= 0 || (!hasHeaders && entry.rows.length === 0)) {
    return { block: '', height: 0 };
  }

  const availableWidth = CONTEXT_TABLE_WIDTH - (CONTEXT_CELL_MARGIN * 2);

  const measureTextWeight = (text) => {
    const normalized = normalizeContextTableCell(text);
    if (!normalized) {
      return 1;
    }

    const lines = normalized.replace(/\r\n/g, '\n').split('\n');
    let maxLength = 1;
    lines.forEach((line) => {
      const length = line.trim().length;
      if (length > 0) {
        maxLength = Math.max(maxLength, length);
      } else {
        maxLength = Math.max(maxLength, 1);
      }
    });

    return Math.max(maxLength, lines.length * 4);
  };

  const columnWeights = Array.from({ length: columnCount }, (_, columnIndex) => {
    const headerWeight = hasHeaders ? measureTextWeight(entry.headers[columnIndex]) : 1;
    const rowWeight = entry.rows.reduce((max, row) => {
      const cell = row[columnIndex] ?? '';
      return Math.max(max, measureTextWeight(cell));
    }, 1);

    return Math.max(1, headerWeight, rowWeight);
  });

  const totalWeight = columnWeights.reduce((sum, weight) => sum + weight, 0) || columnCount;
  const minimumWidth = Math.max(1200, Math.floor(availableWidth / (columnCount * 2)));

  const columnWidths = columnWeights.map((weight) => {
    const proportionalWidth = Math.floor((weight / totalWeight) * availableWidth);
    return Math.max(minimumWidth, proportionalWidth);
  });

  let widthSum = columnWidths.reduce((sum, width) => sum + width, 0);
  let widthDelta = availableWidth - widthSum;

  if (widthDelta !== 0 && columnWidths.length > 0) {
    if (widthDelta > 0) {
      columnWidths[columnWidths.length - 1] += widthDelta;
    } else {
      let remaining = widthDelta;
      for (let index = columnWidths.length - 1; index >= 0 && remaining < 0; index -= 1) {
        const currentWidth = columnWidths[index];
        const adjustable = currentWidth - minimumWidth;
        if (adjustable <= 0) {
          continue;
        }

        const adjustment = Math.max(remaining, -adjustable);
        columnWidths[index] = currentWidth + adjustment;
        remaining -= adjustment;
      }

      if (remaining !== 0) {
        columnWidths[columnWidths.length - 1] = Math.max(minimumWidth, columnWidths[columnWidths.length - 1] + remaining);
      }
    }

    widthSum = columnWidths.reduce((sum, width) => sum + width, 0);
    if (widthSum !== availableWidth) {
      columnWidths[columnWidths.length - 1] += availableWidth - widthSum;
    }
  }

  const rowHeights = [];
  const tableRows = [];

  if (hasHeaders) {
    const headerHeight = estimateDataTableRowHeight(entry.headers);
    rowHeights.push(headerHeight);

    const headerCells = entry.headers.map((headerText, colIndex) => {
      const cellParagraphId = paragraphIdFactory();
      const usableWidth = Math.max(1200, columnWidths[colIndex] - (NESTED_IN_MARGIN.left + NESTED_IN_MARGIN.right));
      const paragraph = buildParagraph({
        id: cellParagraphId,
        paraPrIDRef: '62',
        styleIDRef: '2',
        charPrIDRef: '63',
        text: headerText,
        lineSegOptions: {
          baseline: 808,
          spacing: 188,
          horzsize: Math.max(1200, usableWidth),
          flags: '393216'
        },
        indent: '              '
      });

      return [
        '          <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="3">',
        '            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">',
        paragraph,
        '            </hp:subList>',
        `            <hp:cellAddr colAddr="${colIndex}" rowAddr="0" />`,
        '            <hp:cellSpan colSpan="1" rowSpan="1" />',
        `            <hp:cellSz width="${columnWidths[colIndex]}" height="${headerHeight}" />`,
        `            <hp:cellMargin left="${NESTED_IN_MARGIN.left}" right="${NESTED_IN_MARGIN.right}" top="${NESTED_IN_MARGIN.top}" bottom="${NESTED_IN_MARGIN.bottom}" />`,
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
      const usableWidth = Math.max(1200, columnWidths[colIndex] - (NESTED_IN_MARGIN.left + NESTED_IN_MARGIN.right));
      const paragraph = buildParagraph({
        id: cellParagraphId,
        paraPrIDRef: '44',
        styleIDRef: '46',
        charPrIDRef: '65',
        text: cellText,
        lineSegOptions: {
          baseline: 808,
          spacing: 380,
          horzsize: Math.max(1200, usableWidth),
          flags: '393216'
        },
        indent: '              '
      });

      return [
        '          <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="3">',
        '            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">',
        paragraph,
        '            </hp:subList>',
        `            <hp:cellAddr colAddr="${colIndex}" rowAddr="${rowAddress}" />`,
        '            <hp:cellSpan colSpan="1" rowSpan="1" />',
        `            <hp:cellSz width="${columnWidths[colIndex]}" height="${rowHeight}" />`,
        `            <hp:cellMargin left="${NESTED_IN_MARGIN.left}" right="${NESTED_IN_MARGIN.right}" top="${NESTED_IN_MARGIN.top}" bottom="${NESTED_IN_MARGIN.bottom}" />`,
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
    return { block: '', height: 0 };
  }

  const tableHeight = rowHeights.reduce((total, height) => total + height, 1415);
  const bottomMargin = hasFollowingEntry ? NESTED_TABLE_OUT_MARGIN_BOTTOM : 0;
  const tableId = tableIdFactory();

  return {
    height: tableHeight,
    block: [
      `                  <hp:tbl id="${tableId}" zOrder="12" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="1" rowCnt="${tableRows.length}" colCnt="${columnCount}" cellSpacing="0" borderFillIDRef="3" noAdjust="0">`,
      `                    <hp:sz width="${CONTEXT_TABLE_WIDTH}" widthRelTo="ABSOLUTE" height="${tableHeight}" heightRelTo="ABSOLUTE" protect="0" />`,
      '                    <hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0" />',
      `                    <hp:outMargin left="0" right="0" top="${NESTED_TABLE_OUT_MARGIN_TOP}" bottom="${bottomMargin}" />`,
      `                    <hp:inMargin left="${NESTED_IN_MARGIN.left}" right="${NESTED_IN_MARGIN.right}" top="${NESTED_IN_MARGIN.top}" bottom="${NESTED_IN_MARGIN.bottom}" />`,
      tableRows.join('\n'),
      '                  </hp:tbl>'
    ].join('\n')
  };
}

function buildContextRow({
  entry,
  rowIndex,
  paragraphIdFactory,
  tableIdFactory,
  hasFollowingEntry
}) {
  const cellBlocks = [];
  let rowHeight = 0;

  if (entry.type === 'text') {
    const displayText = formatDisplayText(entry);
    const paragraphId = paragraphIdFactory();
    const paragraph = buildParagraph({
      id: paragraphId,
      paraPrIDRef: '44',
      styleIDRef: '0',
      charPrIDRef: '53',
      text: displayText,
      lineSegOptions: {
        spacing: 460,
        horzsize: 28908,
        flags: '393216'
      },
      indent: '              '
    });
    cellBlocks.push(paragraph);
  rowHeight += estimateContextRowHeight(displayText) + CONTEXT_ROW_EXTRA_PADDING;
  } else if (entry.type === 'table') {
    const descriptionText = entry.displayText ?? formatDisplayText(entry);

    const tableBlock = buildContextDataTable({
      entry,
      paragraphIdFactory,
      tableIdFactory,
      hasFollowingEntry
    });

    const descriptionHeight = descriptionText ? estimateContextRowHeight(descriptionText) : 0;

    if (descriptionText) {
      const descriptionParagraphId = paragraphIdFactory();
      const descriptionParagraph = buildParagraph({
        id: descriptionParagraphId,
        paraPrIDRef: '44',
        styleIDRef: '0',
        charPrIDRef: NESTED_PARAGRAPH_CHAR_PR_ID,
        text: descriptionText,
        lineSegOptions: {
          spacing: 460,
          horzsize: 28908,
          flags: '393216'
        },
        indent: '              '
      });
      cellBlocks.push(descriptionParagraph);
      rowHeight += descriptionHeight;
    }

    const tableParagraphHeight = CONTEXT_ROW_EXTRA_PADDING + (tableBlock.height || 0);

    if (tableBlock.block) {
      const tableParagraphId = paragraphIdFactory();
      const tableLineseg = buildLinesegArray({
        vertsize: tableParagraphHeight,
        textheight: tableParagraphHeight,
        baseline: Math.max(978, tableParagraphHeight - 360),
        spacing: 460,
        horzpos: 0,
        horzsize: 28908,
        flags: '393216'
      }, '                ');

      cellBlocks.push([
        `              <hp:p id="${tableParagraphId}" paraPrIDRef="${NESTED_PARAGRAPH_PARA_PR_ID}" styleIDRef="${NESTED_PARAGRAPH_STYLE_ID}" pageBreak="0" columnBreak="0" merged="0">`,
        `                <hp:run charPrIDRef="${NESTED_PARAGRAPH_CHAR_PR_ID}">`,
        tableBlock.block,
        '                  <hp:t />',
        '                </hp:run>',
        tableLineseg,
        '              </hp:p>'
      ].join('\n'));
    }

    rowHeight += tableParagraphHeight;
  }

  if (cellBlocks.length === 0) {
    return { row: '', height: 0 };
  }

  const cellString = [
    '        <hp:tr>',
    `          <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="${CONTEXT_CELL_BORDER_ID}">`,
    '            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">',
    cellBlocks.join('\n'),
    '            </hp:subList>',
    `            <hp:cellAddr colAddr="0" rowAddr="${rowIndex}" />`,
    '            <hp:cellSpan colSpan="1" rowSpan="1" />',
    `            <hp:cellSz width="${CONTEXT_TABLE_WIDTH}" height="${rowHeight}" />`,
    `            <hp:cellMargin left="${CONTEXT_CELL_MARGIN}" right="${CONTEXT_CELL_MARGIN}" top="${CONTEXT_CELL_MARGIN}" bottom="${CONTEXT_CELL_MARGIN}" />`,
    '          </hp:tc>',
    '        </hp:tr>'
  ].join('\n');

  return { row: cellString, height: rowHeight };
}

function buildContextTable({
  paragraphId,
  tableId,
  entries,
  paragraphIdFactory,
  tableIdFactory
}) {
  if (!entries || entries.length === 0) {
    return '';
  }

  const rows = [];
  const rowHeights = [];

  entries.forEach((entry, index) => {
    const result = buildContextRow({
      entry,
      rowIndex: index,
      paragraphIdFactory,
      tableIdFactory,
      hasFollowingEntry: index < entries.length - 1
    });
    if (result.row) {
      rows.push(result.row);
      rowHeights.push(result.height || 0);
    }
  });

  if (rows.length === 0) {
    return '';
  }

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
    '    <hp:run charPrIDRef="60">',
    `      <hp:tbl id="${tableId}" zOrder="12" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="1" rowCnt="${rows.length}" colCnt="1" cellSpacing="0" borderFillIDRef="${CONTEXT_OUTER_BORDER_ID}" noAdjust="0">`,
    `        <hp:sz width="${CONTEXT_TABLE_WIDTH}" widthRelTo="ABSOLUTE" height="${tableHeight}" heightRelTo="ABSOLUTE" protect="0" />`,
    '        <hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0" />',
    '        <hp:outMargin left="0" right="0" top="0" bottom="0" />',
    `        <hp:inMargin left="${CONTEXT_CELL_MARGIN}" right="${CONTEXT_CELL_MARGIN}" top="${CONTEXT_CELL_MARGIN}" bottom="${CONTEXT_CELL_MARGIN}" />`,
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
  const mergedEntries = mergeContextEntries(normalizedEntries);
  const tableParagraphId = paragraphIdFactory();
  const tableId = tableIdFactory();
  const table = buildContextTable({
    paragraphId: tableParagraphId,
    tableId,
    entries: mergedEntries,
    paragraphIdFactory,
    tableIdFactory
  });

  return table ? [table] : [];
}

export {
  buildContextBlocks,
  estimateContextRowHeight,
  normalizeContextEntry
};
