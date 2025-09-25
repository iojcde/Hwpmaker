import { buildParagraph, buildLinesegArray } from './paragraphs.js';
import { DEFAULT_CHOICE_NUMERALS } from '../constants.js';

function normalizeChoice(choice) {
  if (choice && typeof choice === 'object') {
    return choice.text ?? choice.content ?? choice.description ?? '';
  }

  return choice === undefined || choice === null ? '' : String(choice);
}

function formatChoiceText(choice, index, options) {
  const choiceNumerals = options.choiceNumerals ?? DEFAULT_CHOICE_NUMERALS;
  const numeral = choiceNumerals[index] ?? `${index + 1}`;
  return `${numeral} ${choice}`;
}

function buildChoiceTable({
  paragraphId,
  tableId,
  choices,
  options,
  paragraphIdFactory
}) {
  if (!Array.isArray(choices) || choices.length === 0) {
    return '';
  }

  const normalizedChoices = choices.map(normalizeChoice);
  const colCount = Math.max(1, normalizedChoices.length * 2);
  const rowCount = Math.ceil(normalizedChoices.length / colCount);

  const cellWidth = Math.floor(30611 / colCount);
  const lastCellWidth = 30611 - cellWidth * (colCount - 1);

  const tableRows = [];
  for (let row = 0; row < rowCount; row += 1) {
    const cells = [];
    for (let col = 0; col < colCount; col += 1) {
      const choiceIndex = row * colCount + col;
      if (choiceIndex < normalizedChoices.length) {
        const choice = normalizedChoices[choiceIndex];
        const choiceText = formatChoiceText(choice, choiceIndex, options);
        const cellParagraphId = paragraphIdFactory();
        const paragraph = buildParagraph({
          id: cellParagraphId,
          paraPrIDRef: '60',
          styleIDRef: '44',
          charPrIDRef: '46',
          text: choiceText,
          lineSegOptions: {
            spacing: 516,
            horzsize: Math.max(1800, (col === colCount - 1 ? lastCellWidth : cellWidth) - 1500),
            flags: '1441792'
          },
          indent: '              '
        });

        cells.push([
          '          <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="6">',
          '            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">',
          paragraph,
          '            </hp:subList>',
          `            <hp:cellAddr colAddr="${col}" rowAddr="${row}" />`,
          '            <hp:cellSpan colSpan="1" rowSpan="1" />',
          `            <hp:cellSz width="${col === colCount - 1 ? lastCellWidth : cellWidth}" height="1150" />`,
          '            <hp:cellMargin left="850" right="850" top="850" bottom="850" />',
          '          </hp:tc>'
        ].join('\n'));
      }
    }

    if (cells.length > 0) {
      tableRows.push([
        '        <hp:tr>',
        cells.join('\n'),
        '        </hp:tr>'
      ].join('\n'));
    }
  }

  if (tableRows.length === 0) {
    return '';
  }

  const tableHeight = rowCount * 1150 + 1700;
  const linesegArray = buildLinesegArray({
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
    `      <hp:tbl id="${tableId}" zOrder="12" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="1" rowCnt="${rowCount}" colCnt="${colCount}" cellSpacing="0" borderFillIDRef="22" noAdjust="0">`,
    '        <hp:sz width="30611" widthRelTo="ABSOLUTE" height="' + tableHeight + '" heightRelTo="ABSOLUTE" protect="0" />',
    '        <hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0" />',
    '        <hp:outMargin left="0" right="0" top="0" bottom="0" />',
    '        <hp:inMargin left="850" right="850" top="850" bottom="850" />',
    tableRows.join('\n'),
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
  normalizeChoice,
  formatChoiceText,
  buildChoiceTable
};