import { escapeXml } from '../utils.js';

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
  return [
    `${indent}<hp:linesegarray>`,
    `${indent}  <hp:lineseg textpos="${textpos}" vertpos="${vertpos}" vertsize="${vertsize}" textheight="${textheight}" baseline="${baseline}" spacing="${spacing}" horzpos="${horzpos}" horzsize="${horzsize}" flags="${flags}" />`,
    `${indent}</hp:linesegarray>`
  ].join('\n');
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
  const escapedText = escapeXml(text);
  const linesegArray = buildLinesegArray(lineSegOptions, `${indent}  `);

  const runContent = escapedText
    ? [`${indent}    <hp:t>${escapedText}</hp:t>`]
    : includeEmptyRun ? [] : [`${indent}    <hp:t />`];

  return [
    `${indent}<hp:p id="${id}" paraPrIDRef="${paraPrIDRef}" styleIDRef="${styleIDRef}" pageBreak="0" columnBreak="0" merged="0">`,
    `${indent}  <hp:run charPrIDRef="${charPrIDRef}">`,
    ...runContent,
    `${indent}  </hp:run>`,
    linesegArray,
    `${indent}</hp:p>`
  ].join('\n');
}

function buildSectionOpeningParagraph({ prompt, paragraphId, tableId }) {
  const escapedPrompt = escapeXml(prompt);
  return `  <hp:p id="${paragraphId}" paraPrIDRef="58" styleIDRef="1" pageBreak="0" columnBreak="0" merged="0">
    <hp:run charPrIDRef="3">
      <hp:ctrl>
        <hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="2" sameSz="1" sameGap="3120" />
      </hp:ctrl>
      <hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="1" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="4">
        <hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0" />
        <hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0" />
        <hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0" />
        <hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0" />
        <hp:pagePr landscape="WIDELY" width="77102" height="111685" gutterType="LEFT_RIGHT">
          <hp:margin header="4960" footer="3401" gutter="0" left="5300" right="5300" top="6236" bottom="5952" />
        </hp:pagePr>
        <hp:footNotePr>
          <hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0" />
          <hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000" />
          <hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850" />
          <hp:numbering type="CONTINUOUS" newNum="1" />
          <hp:placement place="EACH_COLUMN" beneathText="0" />
        </hp:footNotePr>
        <hp:endNotePr>
          <hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0" />
          <hp:noteLine length="14692344" type="SOLID" width="0.12 mm" color="#000000" />
          <hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850" />
          <hp:numbering type="CONTINUOUS" newNum="1" />
          <hp:placement place="END_OF_DOCUMENT" beneathText="0" />
        </hp:endNotePr>
        <hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">
          <hp:offset left="1417" right="1417" top="1417" bottom="1417" />
        </hp:pageBorderFill>
        <hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">
          <hp:offset left="1417" right="1417" top="1417" bottom="1417" />
        </hp:pageBorderFill>
        <hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">
          <hp:offset left="1417" right="1417" top="1417" bottom="1417" />
        </hp:pageBorderFill>
        <hp:masterPage idRef="masterpage0" />
        <hp:masterPage idRef="masterpage1" />
        <hp:masterPage idRef="masterpage2" />
        <hp:masterPage idRef="masterpage3" />
      </hp:secPr>
    </hp:run>
    <hp:run charPrIDRef="3">
      <hp:ctrl>
        <hp:newNum num="1" numType="PAGE" />
      </hp:ctrl>
      <hp:ctrl>
        <hp:footer id="3" applyPageType="BOTH">
          <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="BOTTOM" linkListIDRef="0" linkListNextIDRef="0" textWidth="66502" textHeight="3401" hasTextRef="0" hasNumRef="0">
            <hp:p id="2147483648" paraPrIDRef="56" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
              <hp:run charPrIDRef="54" />
              <hp:linesegarray>
                <hp:lineseg textpos="0" vertpos="0" vertsize="1150" textheight="1150" baseline="1150" spacing="460" horzpos="0" horzsize="66500" flags="393216" />
              </hp:linesegarray>
            </hp:p>
          </hp:subList>
        </hp:footer>
      </hp:ctrl>
      <hp:ctrl>
        <hp:header id="1" applyPageType="ODD">
          <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="66502" textHeight="4960" hasTextRef="0" hasNumRef="0">
            <hp:p id="0" paraPrIDRef="45" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
              <hp:run charPrIDRef="54" />
              <hp:linesegarray>
                <hp:lineseg textpos="0" vertpos="0" vertsize="1150" textheight="1150" baseline="978" spacing="460" horzpos="0" horzsize="66500" flags="393216" />
              </hp:linesegarray>
            </hp:p>
          </hp:subList>
        </hp:header>
      </hp:ctrl>
      <hp:ctrl>
        <hp:header id="2" applyPageType="EVEN">
          <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="66502" textHeight="4960" hasTextRef="0" hasNumRef="0">
            <hp:p id="0" paraPrIDRef="28" styleIDRef="33" pageBreak="0" columnBreak="0" merged="0">
              <hp:run charPrIDRef="25" />
              <hp:linesegarray>
                <hp:lineseg textpos="0" vertpos="0" vertsize="900" textheight="900" baseline="765" spacing="452" horzpos="0" horzsize="66500" flags="393216" />
              </hp:linesegarray>
            </hp:p>
          </hp:subList>
        </hp:header>
      </hp:ctrl>
      <hp:tbl id="${tableId}" zOrder="8" numberingType="TABLE" textWrap="SQUARE" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="1" rowCnt="1" colCnt="1" cellSpacing="0" borderFillIDRef="3" noAdjust="0">
        <hp:sz width="66472" widthRelTo="ABSOLUTE" height="13888" heightRelTo="ABSOLUTE" protect="0" />
        <hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PAPER" horzRelTo="PAGE" vertAlign="TOP" horzAlign="CENTER" vertOffset="5215" horzOffset="0" />
        <hp:outMargin left="0" right="0" top="0" bottom="1134" />
        <hp:inMargin left="0" right="0" top="0" bottom="0" />
        <hp:tr>
          <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="20">
            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
              <hp:p id="2147483648" paraPrIDRef="47" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                <hp:run charPrIDRef="53" />
                <hp:linesegarray>
                  <hp:lineseg textpos="0" vertpos="0" vertsize="1150" textheight="1150" baseline="978" spacing="804" horzpos="0" horzsize="66472" flags="393216" />
                </hp:linesegarray>
              </hp:p>
            </hp:subList>
            <hp:cellAddr colAddr="0" rowAddr="0" />
            <hp:cellSpan colSpan="1" rowSpan="1" />
            <hp:cellSz width="66472" height="13888" />
            <hp:cellMargin left="1701" right="1701" top="1701" bottom="1701" />
          </hp:tc>
        </hp:tr>
      </hp:tbl>
    </hp:run>
    <hp:run charPrIDRef="3">
      <hp:t>${escapedPrompt}</hp:t>
    </hp:run>
    <hp:linesegarray>
      <hp:lineseg textpos="0" vertpos="0" vertsize="1150" textheight="1150" baseline="978" spacing="460" horzpos="0" horzsize="31688" flags="393216" />
    </hp:linesegarray>
  </hp:p>`;
}

function buildPromptParagraph({ prompt, paragraphId }) {
  return buildParagraph({
    id: paragraphId,
    paraPrIDRef: '46',
    styleIDRef: '0',
    charPrIDRef: '49',
    text: prompt,
    lineSegOptions: {
      horzsize: 31688
    }
  });
}

function buildChoiceParagraph({ paragraphId, text }) {
  return buildParagraph({
    id: paragraphId,
    paraPrIDRef: '46',
    styleIDRef: '0',
    charPrIDRef: '49',
    text,
    lineSegOptions: {
      horzsize: 31688
    }
  });
}

function buildSpacerParagraph(paragraphId) {
  return buildParagraph({
    id: paragraphId,
    paraPrIDRef: '46',
    styleIDRef: '0',
    charPrIDRef: '49',
    text: '',
    lineSegOptions: {
      horzsize: 31688
    },
    includeEmptyRun: true
  });
}

export {
  buildLinesegArray,
  buildParagraph,
  buildSectionOpeningParagraph,
  buildPromptParagraph,
  buildChoiceParagraph,
  buildSpacerParagraph
};