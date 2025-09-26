import { DEFAULT_CHOICE_NUMERALS } from './constants.js';
import { buildParagraph, buildLinesegArray } from './paragraph.js';

function normalizeChoice(choice) {
  if (choice && typeof choice === 'object') {
    return choice.text ?? choice.content ?? choice.description ?? '';
  }

  return choice === undefined || choice === null ? '' : String(choice);
}

function formatChoiceText(choice, index, options) {
  const numerals = options.choiceNumerals ?? DEFAULT_CHOICE_NUMERALS;
  const prefix = numerals[index] ?? `${index + 1}.`;
  return `${prefix} ${choice}`;
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
      paraPrIDRef: '44',
      styleIDRef: '0',
      charPrIDRef: '53',
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
      '          <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="22">',
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
      paraPrIDRef: '44',
      styleIDRef: '0',
      charPrIDRef: '53',
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
      '          <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="22">',
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
    `  <hp:p id="${paragraphId}" paraPrIDRef="60" styleIDRef="10" pageBreak="0" columnBreak="0" merged="0">`,
    '    <hp:run charPrIDRef="0">',
    `      <hp:tbl id="${tableId}" zOrder="21" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="1" rowCnt="1" colCnt="${columnCount}" cellSpacing="0" borderFillIDRef="5" noAdjust="0">`,
    '        <hp:sz width="30615" widthRelTo="ABSOLUTE" height="1431" heightRelTo="ABSOLUTE" protect="1" />',
    '        <hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0" />',
    '        <hp:outMargin left="0" right="0" top="0" bottom="0" />',
    '        <hp:inMargin left="0" right="0" top="0" bottom="0" />',
    '        <hp:tr>',
    cells.join('\n'),
    '        </hp:tr>',
    '      </hp:tbl>',
    '    </hp:run>',
    '    <hp:run charPrIDRef="64">',
    '      <hp:t />',
    '    </hp:run>',
    lineseg,
    '  </hp:p>'
  ].join('\n');
}

export {
  normalizeChoice,
  formatChoiceText,
  buildChoiceParagraph,
  buildSpacerParagraph,
  buildChoiceTable
};
