import { escapeXml } from './utils.js';

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
  return `  <hp:p id="${paragraphId}" paraPrIDRef="55" styleIDRef="1" pageBreak="0" columnBreak="0" merged="0">\n    <hp:run charPrIDRef="3">\n      <hp:ctrl>\n        <hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="2" sameSz="1" sameGap="3120" />\n      </hp:ctrl>\n      <hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="1" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="4">\n        <hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0" />\n        <hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0" />\n        <hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0" />\n        <hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0" />\n        <hp:pagePr landscape="WIDELY" width="77102" height="111685" gutterType="LEFT_RIGHT">\n          <hp:margin header="4960" footer="3401" gutter="0" left="5300" right="5300" top="6236" bottom="5952" />\n        </hp:pagePr>\n        <hp:footNotePr>\n          <hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0" />\n          <hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000" />\n          <hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850" />\n          <hp:numbering type="CONTINUOUS" newNum="1" />\n          <hp:placement place="EACH_COLUMN" beneathText="0" />\n        </hp:footNotePr>\n        <hp:endNotePr>\n          <hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0" />\n          <hp:noteLine length="14692344" type="SOLID" width="0.12 mm" color="#000000" />\n          <hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850" />\n          <hp:numbering type="CONTINUOUS" newNum="1" />\n          <hp:placement place="END_OF_DOCUMENT" beneathText="0" />\n        </hp:endNotePr>\n        <hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">\n          <hp:offset left="1417" right="1417" top="1417" bottom="1417" />\n        </hp:pageBorderFill>\n        <hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">\n          <hp:offset left="1417" right="1417" top="1417" bottom="1417" />\n        </hp:pageBorderFill>\n        <hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">\n          <hp:offset left="1417" right="1417" top="1417" bottom="1417" />\n        </hp:pageBorderFill>\n        <hp:masterPage idRef="masterpage0" />\n        <hp:masterPage idRef="masterpage1" />\n        <hp:masterPage idRef="masterpage2" />\n        <hp:masterPage idRef="masterpage3" />\n      </hp:secPr>\n    </hp:run>\n    <hp:run charPrIDRef="3">\n      <hp:ctrl>\n        <hp:newNum num="1" numType="PAGE" />\n      </hp:ctrl>\n      <hp:ctrl>\n        <hp:footer id="3" applyPageType="BOTH">\n          <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="BOTTOM" linkListIDRef="0" linkListNextIDRef="0" textWidth="66502" textHeight="3401" hasTextRef="0" hasNumRef="0">\n            <hp:p id="2147483648" paraPrIDRef="54" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">\n              <hp:run charPrIDRef="53" />\n              <hp:linesegarray>\n                <hp:lineseg textpos="0" vertpos="0" vertsize="1150" textheight="1150" baseline="1150" spacing="460" horzpos="0" horzsize="66500" flags="393216" />\n              </hp:linesegarray>\n            </hp:p>\n          </hp:subList>\n        </hp:footer>\n      </hp:ctrl>\n      <hp:ctrl>\n        <hp:header id="1" applyPageType="ODD">\n          <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="66502" textHeight="4960" hasTextRef="0" hasNumRef="0">\n            <hp:p id="0" paraPrIDRef="44" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">\n              <hp:run charPrIDRef="53" />\n              <hp:linesegarray>\n                <hp:lineseg textpos="0" vertpos="0" vertsize="1150" textheight="1150" baseline="978" spacing="460" horzpos="0" horzsize="66500" flags="393216" />\n              </hp:linesegarray>\n            </hp:p>\n          </hp:subList>\n        </hp:header>\n      </hp:ctrl>\n      <hp:ctrl>\n        <hp:header id="2" applyPageType="EVEN">\n          <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="66502" textHeight="4960" hasTextRef="0" hasNumRef="0">\n            <hp:p id="0" paraPrIDRef="28" styleIDRef="33" pageBreak="0" columnBreak="0" merged="0">\n              <hp:run charPrIDRef="25" />\n              <hp:linesegarray>\n                <hp:lineseg textpos="0" vertpos="0" vertsize="900" textheight="900" baseline="765" spacing="452" horzpos="0" horzsize="66500" flags="393216" />\n              </hp:linesegarray>\n            </hp:p>\n          </hp:subList>\n        </hp:header>\n      </hp:ctrl>\n      <hp:tbl id="${tableId}" zOrder="8" numberingType="TABLE" textWrap="SQUARE" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="1" rowCnt="1" colCnt="1" cellSpacing="0" borderFillIDRef="3" noAdjust="0">\n        <hp:sz width="66472" widthRelTo="ABSOLUTE" height="13888" heightRelTo="ABSOLUTE" protect="0" />\n        <hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PAPER" horzRelTo="PAGE" vertAlign="TOP" horzAlign="CENTER" vertOffset="5215" horzOffset="0" />\n        <hp:outMargin left="0" right="0" top="0" bottom="1134" />\n        <hp:inMargin left="0" right="0" top="0" bottom="0" />\n        <hp:tr>\n          <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="20">\n            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">\n              <hp:p id="2147483648" paraPrIDRef="45" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">\n                <hp:run charPrIDRef="52" />\n                <hp:linesegarray>\n                  <hp:lineseg textpos="0" vertpos="0" vertsize="1150" textheight="1150" baseline="978" spacing="804" horzpos="0" horzsize="66472" flags="393216" />\n                </hp:linesegarray>\n              </hp:p>\n            </hp:subList>\n            <hp:cellAddr colAddr="0" rowAddr="0" />\n            <hp:cellSpan colSpan="1" rowSpan="1" />\n            <hp:cellSz width="66472" height="13888" />\n            <hp:cellMargin left="141" right="141" top="141" bottom="141" />\n          </hp:tc>\n        </hp:tr>\n      </hp:tbl>\n      <hp:ctrl>\n        <hp:pageHiding hideHeader="1" hideFooter="0" hideMasterPage="0" hideBorder="0" hideFill="0" hidePageNum="0" />\n      </hp:ctrl>\n      <hp:t>${escapedPrompt}</hp:t>\n    </hp:run>\n    <hp:linesegarray>\n      <hp:lineseg textpos="0" vertpos="9041" vertsize="1400" textheight="1400" baseline="1190" spacing="560" horzpos="0" horzsize="31688" flags="2490368" />\n      <hp:lineseg textpos="101" vertpos="11001" vertsize="1150" textheight="1150" baseline="978" spacing="460" horzpos="0" horzsize="31688" flags="1441792" />\n    </hp:linesegarray>\n  </hp:p>`;
}

function buildPromptParagraph({ prompt, paragraphId }) {
  return buildParagraph({
    id: paragraphId,
    paraPrIDRef: '55',
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

export {
  buildLinesegArray,
  buildParagraph,
  buildSectionOpeningParagraph,
  buildPromptParagraph
};
