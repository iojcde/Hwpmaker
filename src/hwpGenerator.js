const DEFAULT_BASE_PARAGRAPH_ID = 2147483649;
const DEFAULT_BASE_TABLE_ID = 1900000001;

export const DEFAULT_CHOICE_NUMERALS = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩'];
export const KOREAN_CHOICE_LABELS = ['ㄱ.', 'ㄴ.', 'ㄷ.', 'ㄹ.', 'ㅁ.', 'ㅂ.', 'ㅅ.', 'ㅇ.', 'ㅈ.', 'ㅊ.', 'ㅋ.', 'ㅌ.', 'ㅍ.', 'ㅎ.'];

export const DEFAULT_MINIFY_OPTIONS = {
	removeComments: true,
	removeWhitespaceBetweenTags: true,
	considerPreserveWhitespace: true,
	collapseWhitespaceInTags: true,
	collapseEmptyElements: true,
	trimWhitespaceFromTexts: false,
	collapseWhitespaceInTexts: true,
	collapseWhitespaceInProlog: true,
	collapseWhitespaceInDocType: true,
	removeSchemaLocationAttributes: false,
	removeUnnecessaryStandaloneDeclaration: true,
	removeUnusedNamespaces: true,
	removeUnusedDefaultNamespace: true,
	shortenNamespaces: false,
	ignoreCData: true
};

const DEFAULT_OPTIONS = {
	choiceNumerals: DEFAULT_CHOICE_NUMERALS,
	spacersPerQuestion: 1,
	choiceLayout: null,
	baseParagraphId: DEFAULT_BASE_PARAGRAPH_ID,
	baseTableId: DEFAULT_BASE_TABLE_ID,
	answerHeading: '정답 및 해설',
	answerSpacersPerQuestion: null,
	baseAnswerParagraphId: null,
	baseAnswerTableId: null
};

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8"?>';

function escapeXml(value) {
	if (value === null || value === undefined) {
		return '';
	}

	return String(value)
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&apos;');
}

function decodeCommonEntities(value) {
	if (!value) {
		return '';
	}

	return String(value)
		.replace(/&lt;/gi, '<')
		.replace(/&gt;/gi, '>')
		.replace(/&amp;/gi, '&');
}

function createIdCounter(start) {
	let current = start;
	return () => {
		const next = current;
		current += 1;
		return next;
	};
}

function ensureQuestions(value) {
	if (!value || !Array.isArray(value) || value.length === 0) {
		throw new Error('"questions" must be a non-empty array.');
	}

	value.forEach((question, index) => {
		if (!question || typeof question !== 'object') {
			throw new Error(`Question at index ${index} must be an object.`);
		}

		if (!question.prompt || typeof question.prompt !== 'string' || question.prompt.trim() === '') {
			throw new Error(`Question at index ${index} is missing a prompt.`);
		}

		if (!Array.isArray(question.choices) || question.choices.length === 0) {
			throw new Error(`Question at index ${index} must include at least one choice.`);
		}
	});

	return value;
}

function coerceOptions(options) {
	const normalized = { ...DEFAULT_OPTIONS, ...(options || {}) };

	if (!Array.isArray(normalized.choiceNumerals) || normalized.choiceNumerals.length === 0) {
		normalized.choiceNumerals = DEFAULT_CHOICE_NUMERALS.slice();
	}

	normalized.spacersPerQuestion = Number.isInteger(normalized.spacersPerQuestion)
		? Math.max(0, normalized.spacersPerQuestion)
		: DEFAULT_OPTIONS.spacersPerQuestion;

	normalized.answerSpacersPerQuestion = Number.isInteger(normalized.answerSpacersPerQuestion)
		? Math.max(0, normalized.answerSpacersPerQuestion)
		: normalized.spacersPerQuestion;

	normalized.baseParagraphId = Number.isInteger(normalized.baseParagraphId)
		? normalized.baseParagraphId
		: DEFAULT_BASE_PARAGRAPH_ID;

	normalized.baseTableId = Number.isInteger(normalized.baseTableId)
		? normalized.baseTableId
		: DEFAULT_BASE_TABLE_ID;

	normalized.baseAnswerParagraphId = Number.isInteger(normalized.baseAnswerParagraphId)
		? normalized.baseAnswerParagraphId
		: normalized.baseParagraphId + 10000;

	normalized.baseAnswerTableId = Number.isInteger(normalized.baseAnswerTableId)
		? normalized.baseAnswerTableId
		: normalized.baseTableId + 1000;

	return normalized;
}

function formatStatementTitle(value) {
	if (!value) {
		return '';
	}

	const decoded = decodeCommonEntities(value).trim();
	const match = decoded.match(/^<(.*)>$/);
	if (!match) {
		return escapeXml(decoded);
	}

	const inner = match[1];
	const spaced = inner.split('').join(' ');
	return escapeXml(`<${spaced}>`);
}

function normalizeContextEntries(entries) {
	if (!entries) {
		return [];
	}

	return entries
		.filter((entry) => entry && typeof entry === 'object')
		.map((entry) => {
			const label = entry.label ? String(entry.label) : '';
			const text = entry.text ? String(entry.text) : '';
			const table = entry.table && typeof entry.table === 'object' ? entry.table : null;
			return { label, text, table };
		});
}

function toArray(value) {
	if (!value) {
		return [];
	}

	if (Array.isArray(value)) {
		return value;
	}

	return [value];
}

function extractTableHeaders(table) {
	if (!table || !table.headers) {
		return [];
	}

	return table.headers.map((header, index) => {
		if (typeof header === 'string') {
			return { label: header, key: index };
		}

		if (header && typeof header === 'object') {
			const label = header.label ?? header.key ?? `컬럼 ${index + 1}`;
			const key = header.key ?? index;
			return { label, key };
		}

		return { label: `컬럼 ${index + 1}`, key: index };
	});
}

function extractTableRows(table, headers) {
	if (!table || !table.rows) {
		return [];
	}

	return table.rows.map((row) => {
		if (Array.isArray(row)) {
			return row.map((cell) => (cell === null || cell === undefined ? '' : String(cell)));
		}

		if (row && typeof row === 'object') {
			return headers.map((header, index) => {
				const key = header.key ?? index;
				const value = row[key];
				return value === null || value === undefined ? '' : String(value);
			});
		}

		return headers.map(() => '');
	});
}

function renderLinesegArray({ width = '28908', height = '1150', baseline = '978' } = {}) {
	return `      <hp:linesegarray>
				<hp:lineseg textpos="0" vertpos="0" vertsize="${height}" textheight="${height}" baseline="${baseline}" spacing="460" horzpos="0" horzsize="${width}" flags="393216" />
			</hp:linesegarray>`;
}

function renderPromptParagraph({ paragraphId, prompt, isFirstQuestion, nextTableId }) {
	if (isFirstQuestion) {
		const layoutTableId = nextTableId();
		return `  <hp:p id="${paragraphId}" paraPrIDRef="55" styleIDRef="1" pageBreak="0" columnBreak="0" merged="0">
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
						<hp:p id="2147483648" paraPrIDRef="54" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
							<hp:run charPrIDRef="53" />
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
						<hp:p id="0" paraPrIDRef="44" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
							<hp:run charPrIDRef="53" />
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
			<hp:tbl id="${layoutTableId}" zOrder="8" numberingType="TABLE" textWrap="SQUARE" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="1" rowCnt="1" colCnt="1" cellSpacing="0" borderFillIDRef="3" noAdjust="0">
				<hp:sz width="66472" widthRelTo="ABSOLUTE" height="13888" heightRelTo="ABSOLUTE" protect="0" />
				<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PAPER" horzRelTo="PAGE" vertAlign="TOP" horzAlign="CENTER" vertOffset="5215" horzOffset="0" />
				<hp:outMargin left="0" right="0" top="0" bottom="1134" />
				<hp:inMargin left="0" right="0" top="0" bottom="0" />
				<hp:tr>
					<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="20">
						<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
							<hp:p id="2147483648" paraPrIDRef="45" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
								<hp:run charPrIDRef="52" />
								<hp:linesegarray>
									<hp:lineseg textpos="0" vertpos="0" vertsize="1150" textheight="1150" baseline="978" spacing="804" horzpos="0" horzsize="66472" flags="393216" />
								</hp:linesegarray>
							</hp:p>
						</hp:subList>
						<hp:cellAddr colAddr="0" rowAddr="0" />
						<hp:cellSpan colSpan="1" rowSpan="1" />
						<hp:cellSz width="66472" height="13888" />
						<hp:cellMargin left="141" right="141" top="141" bottom="141" />
					</hp:tc>
				</hp:tr>
			</hp:tbl>
			<hp:ctrl>
				<hp:pageHiding hideHeader="1" hideFooter="0" hideMasterPage="0" hideBorder="0" hideFill="0" hidePageNum="0" />
			</hp:ctrl>
			<hp:t>${escapeXml(prompt)}</hp:t>
		</hp:run>
		<hp:linesegarray>
			<hp:lineseg textpos="0" vertpos="9041" vertsize="1400" textheight="1400" baseline="1190" spacing="560" horzpos="0" horzsize="31688" flags="2490368" />
			<hp:lineseg textpos="101" vertpos="11001" vertsize="1150" textheight="1150" baseline="978" spacing="460" horzpos="0" horzsize="31688" flags="1441792" />
		</hp:linesegarray>
	</hp:p>`;
	}

	return `  <hp:p id="${paragraphId}" paraPrIDRef="55" styleIDRef="1" pageBreak="0" columnBreak="0" merged="0">
		<hp:run charPrIDRef="3">
			<hp:t>${escapeXml(prompt)}</hp:t>
		</hp:run>
${renderLinesegArray({ width: '31688', height: '1400', baseline: '1190' })}
	</hp:p>`;
}

function renderContextTableRow({ entry, nextParagraphId, nextTableId }) {
	const pieces = [];
	const displayedLabel = entry.label ? `${entry.label.trim()} : ` : '';
	const finalText = `${displayedLabel}${entry.text ? entry.text.trim() : ''}`.trim();

	if (finalText) {
		const paragraphId = nextParagraphId();
		pieces.push(`              <hp:p id="${paragraphId}" paraPrIDRef="44" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
								<hp:run charPrIDRef="53">
									<hp:t>${escapeXml(finalText)}</hp:t>
								</hp:run>
${renderLinesegArray({})}
							</hp:p>`);
	}

	if (entry.table) {
		const headers = extractTableHeaders(entry.table);
		const rows = extractTableRows(entry.table, headers);
		const tableId = nextTableId();
		const colCnt = Math.max(headers.length, rows[0] ? rows[0].length : 0, 1);
		const rowCnt = rows.length + (headers.length > 0 ? 1 : 0);

		const headerRow = headers.length === 0
			? ''
			: `                <hp:tr>
${headers.map((header, columnIndex) => `                  <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="3">
										<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
											<hp:p id="${nextParagraphId()}" paraPrIDRef="62" styleIDRef="2" pageBreak="0" columnBreak="0" merged="0">
												<hp:run charPrIDRef="63">
													<hp:t>${escapeXml(header.label)}</hp:t>
												</hp:run>
${renderLinesegArray({ width: '2388', height: '950', baseline: '808' })}
											</hp:p>
										</hp:subList>
										<hp:cellAddr colAddr="${columnIndex}" rowAddr="0" />
										<hp:cellSpan colSpan="1" rowSpan="1" />
										<hp:cellSz width="${Math.round(28910 / Math.max(1, colCnt))}" height="1904" />
										<hp:cellMargin left="141" right="141" top="141" bottom="141" />
									</hp:tc>`).join('\n')}
								</hp:tr>`;

		const dataRows = rows.map((row, rowIndex) => `                <hp:tr>
${row.map((cell, columnIndex) => `                  <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="3">
										<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
											<hp:p id="${nextParagraphId()}" paraPrIDRef="44" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
												<hp:run charPrIDRef="53">
													<hp:t>${escapeXml(cell)}</hp:t>
												</hp:run>
${renderLinesegArray({})}
											</hp:p>
										</hp:subList>
										<hp:cellAddr colAddr="${columnIndex}" rowAddr="${rowIndex + (headers.length ? 1 : 0)}" />
										<hp:cellSpan colSpan="1" rowSpan="1" />
										<hp:cellSz width="${Math.round(28910 / Math.max(1, colCnt))}" height="1904" />
										<hp:cellMargin left="141" right="141" top="141" bottom="141" />
									</hp:tc>`).join('\n')}
								</hp:tr>`).join('\n');

		pieces.push(`              <hp:p id="${nextParagraphId()}" paraPrIDRef="44" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
								<hp:run charPrIDRef="53">
									<hp:tbl id="${tableId}" zOrder="22" numberingType="TABLE" textWrap="SQUARE" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="1" rowCnt="${rowCnt}" colCnt="${colCnt}" cellSpacing="0" borderFillIDRef="3" noAdjust="0">
										<hp:sz width="28910" widthRelTo="ABSOLUTE" height="${Math.max(1, rowCnt) * 1904}" heightRelTo="ABSOLUTE" protect="0" />
										<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0" />
										<hp:outMargin left="0" right="0" top="566" bottom="0" />
										<hp:inMargin left="510" right="510" top="141" bottom="141" />
${headerRow}
${dataRows}
									</hp:tbl>
									<hp:t />
								</hp:run>
${renderLinesegArray({ height: '10344', baseline: '8792' })}
							</hp:p>`);
	}

	return pieces.join('\n');
}

function renderContextTable({ contextEntries, nextParagraphId, nextTableId }) {
	if (!contextEntries.length) {
		return '';
	}

	const tableId = nextTableId();
	const rows = [];

	contextEntries.forEach((entry) => {
		const content = renderContextTableRow({ entry, nextParagraphId, nextTableId });
		const mergeWithPrevious = entry.table && (!entry.text || entry.text.trim() === '') && (!entry.label || entry.label.trim() === '') && rows.length > 0;

		if (mergeWithPrevious) {
			rows[rows.length - 1].content += `\n${content}`;
		} else {
			rows.push({ content });
		}
	});

	const rowCnt = rows.length || 1;
	const paragraphId = nextParagraphId();

	const rowsMarkup = rows.map((row, index) => `        <hp:tr>
					<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="6">
						<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
${row.content}
						</hp:subList>
						<hp:cellAddr colAddr="0" rowAddr="${index}" />
						<hp:cellSpan colSpan="1" rowSpan="1" />
						<hp:cellSz width="30611" height="${Math.max(5640, 21704 / Math.max(1, rowCnt))}" />
						<hp:cellMargin left="850" right="850" top="850" bottom="850" />
					</hp:tc>
				</hp:tr>`).join('\n');

	return `  <hp:p id="${paragraphId}" paraPrIDRef="3" styleIDRef="4" pageBreak="0" columnBreak="0" merged="0">
		<hp:run charPrIDRef="60">
			<hp:tbl id="${tableId}" zOrder="11" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="1" rowCnt="${rowCnt}" colCnt="1" cellSpacing="0" borderFillIDRef="7" noAdjust="0">
				<hp:sz width="30611" widthRelTo="ABSOLUTE" heightRelTo="ABSOLUTE" protect="0" />
				<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0" />
				<hp:outMargin left="0" right="0" top="0" bottom="0" />
				<hp:inMargin left="850" right="850" top="850" bottom="850" />
${rowsMarkup}
			</hp:tbl>
		</hp:run>
		<hp:run charPrIDRef="1">
			<hp:t />
		</hp:run>
${renderLinesegArray({ width: '30558', height: String(Math.max(5640, rowCnt * 5640)), baseline: String(Math.max(18448, rowCnt * 2820)) })}
	</hp:p>`;
}

function renderStatementsTable({ statements, statementTitle, nextParagraphId, nextTableId }) {
	if (!statements.length && !statementTitle) {
		return '';
	}

	const paragraphId = nextParagraphId();
	const tableId = nextTableId();
	const formattedTitle = statementTitle ? formatStatementTitle(statementTitle) : '';
	const statementParagraphs = statements.map((statement, index) => {
		const paragraph = nextParagraphId();
		return `              <hp:p id="${paragraph}" paraPrIDRef="56" styleIDRef="9" pageBreak="0" columnBreak="0" merged="0">
								<hp:run charPrIDRef="53">
									<hp:t>${escapeXml(statement)}</hp:t>
								</hp:run>
${renderLinesegArray({ width: '28856', height: '1150', baseline: '575' })}
							</hp:p>`;
	}).join('\n');

	return `  <hp:p id="${paragraphId}" paraPrIDRef="59" styleIDRef="8" pageBreak="0" columnBreak="0" merged="0">
		<hp:run charPrIDRef="62">
			<hp:tbl id="${tableId}" zOrder="21" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="1" rowCnt="3" colCnt="3" cellSpacing="0" borderFillIDRef="5" noAdjust="0">
				<hp:sz width="30557" widthRelTo="ABSOLUTE" protect="0" />
				<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0" />
				<hp:outMargin left="0" right="0" top="0" bottom="0" />
				<hp:inMargin left="0" right="0" top="0" bottom="0" />
				<hp:tr>
					<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="8">
						<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
							<hp:p id="${nextParagraphId()}" paraPrIDRef="48" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
								<hp:run charPrIDRef="59" />
${renderLinesegArray({ width: '13220', height: '100', baseline: '85' })}
							</hp:p>
						</hp:subList>
						<hp:cellAddr colAddr="0" rowAddr="0" />
						<hp:cellSpan colSpan="1" rowSpan="1" />
						<hp:cellSz width="13223" height="645" />
						<hp:cellMargin left="141" right="141" top="141" bottom="141" />
					</hp:tc>
					<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="9">
						<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
							<hp:p id="${nextParagraphId()}" paraPrIDRef="61" styleIDRef="38" pageBreak="0" columnBreak="0" merged="0">
								<hp:run charPrIDRef="35">
									<hp:t>${formattedTitle}</hp:t>
								</hp:run>
${renderLinesegArray({ width: '4532', height: '1150', baseline: '978' })}
							</hp:p>
						</hp:subList>
						<hp:cellAddr colAddr="1" rowAddr="0" />
						<hp:cellSpan colSpan="1" rowSpan="2" />
						<hp:cellSz width="4535" height="1290" />
						<hp:cellMargin left="141" right="141" top="141" bottom="141" />
					</hp:tc>
					<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="8">
						<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
							<hp:p id="${nextParagraphId()}" paraPrIDRef="48" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
								<hp:run charPrIDRef="59" />
${renderLinesegArray({ width: '12796', height: '100', baseline: '85' })}
							</hp:p>
						</hp:subList>
						<hp:cellAddr colAddr="2" rowAddr="0" />
						<hp:cellSpan colSpan="1" rowSpan="1" />
						<hp:cellSz width="12035" height="645" />
						<hp:cellMargin left="141" right="141" top="141" bottom="141" />
					</hp:tc>
				</hp:tr>
				<hp:tr>
					<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="10">
						<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
							<hp:p id="${nextParagraphId()}" paraPrIDRef="63" styleIDRef="9" pageBreak="0" columnBreak="0" merged="0">
								<hp:run charPrIDRef="59" />
${renderLinesegArray({ width: '13220', height: '100', baseline: '85' })}
							</hp:p>
						</hp:subList>
						<hp:cellAddr colAddr="0" rowAddr="1" />
						<hp:cellSpan colSpan="1" rowSpan="1" />
						<hp:cellSz width="13223" height="645" />
						<hp:cellMargin left="141" right="141" top="141" bottom="141" />
					</hp:tc>
					<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="11">
						<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
							<hp:p id="${nextParagraphId()}" paraPrIDRef="48" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
								<hp:run charPrIDRef="59" />
${renderLinesegArray({ width: '12796', height: '100', baseline: '85' })}
							</hp:p>
						</hp:subList>
						<hp:cellAddr colAddr="2" rowAddr="1" />
						<hp:cellSpan colSpan="1" rowSpan="1" />
						<hp:cellSz width="12035" height="645" />
						<hp:cellMargin left="141" right="141" top="141" bottom="141" />
					</hp:tc>
				</hp:tr>
				<hp:tr>
					<hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="12">
						<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
${statementParagraphs || `              <hp:p id="${nextParagraphId()}" paraPrIDRef="63" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
								<hp:run charPrIDRef="53">
									<hp:t />
								</hp:run>
${renderLinesegArray({ width: '28856', height: '1150', baseline: '575' })}
							</hp:p>`}
						</hp:subList>
						<hp:cellAddr colAddr="0" rowAddr="2" />
						<hp:cellSpan colSpan="3" rowSpan="1" />
						<hp:cellSz width="30557" height="3400" />
						<hp:cellMargin left="850" right="850" top="708" bottom="850" />
					</hp:tc>
				</hp:tr>
			</hp:tbl>
			<hp:t />
		</hp:run>
${renderLinesegArray({ width: '30588', height: '10438', baseline: '8872' })}
	</hp:p>`;
}

function hasKoreanLetters(text) {
	return /[ㄱ-ㅎㅏ-ㅣ가-힣]/.test(text);
}

function generateChoiceNumeral(choice, index, choiceNumerals) {
	// If the choice is exactly a single Korean consonant, use Korean labels
	if (/^[ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ]$/.test(choice.trim())) {
		const koreanIndex = 'ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ'.indexOf(choice.trim());
		return KOREAN_CHOICE_LABELS[koreanIndex] || `${index + 1}.`;
	}
	
	// If there's an explicit numeral provided, use it
	if (choiceNumerals && choiceNumerals[index]) {
		return choiceNumerals[index];
	}
	
	// Default fallback to numbered labels
	return `${index + 1}.`;
}

function chunkChoices(choices, size) {
	const result = [];
	for (let i = 0; i < choices.length; i += size) {
		result.push(choices.slice(i, i + size));
	}
	return result;
}

function renderChoiceTableCells({ numeral, text, nextParagraphId, columnOffset, rowIndex }) {
	const numeralParagraph = nextParagraphId();
	const textParagraph = nextParagraphId();

	return `          <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="22">
						<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
							<hp:p id="${numeralParagraph}" paraPrIDRef="44" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
								<hp:run charPrIDRef="53">
									<hp:t>${escapeXml(numeral)}</hp:t>
								</hp:run>
${renderLinesegArray({ width: '1440', height: '1150', baseline: '978' })}
							</hp:p>
						</hp:subList>
						<hp:cellAddr colAddr="${columnOffset}" rowAddr="${rowIndex}" />
						<hp:cellSpan colSpan="1" rowSpan="1" />
						<hp:cellSz width="1379" height="1431" />
						<hp:cellMargin left="0" right="0" top="0" bottom="0" />
					</hp:tc>
					<hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="22">
						<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
							<hp:p id="${textParagraph}" paraPrIDRef="44" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
								<hp:run charPrIDRef="53">
									<hp:t>${escapeXml(text)}</hp:t>
								</hp:run>
${renderLinesegArray({ width: '4560', height: '1150', baseline: '978' })}
							</hp:p>
						</hp:subList>
						<hp:cellAddr colAddr="${columnOffset + 1}" rowAddr="${rowIndex}" />
						<hp:cellSpan colSpan="1" rowSpan="1" />
						<hp:cellSz width="4744" height="1431" />
						<hp:cellMargin left="184" right="0" top="0" bottom="0" />
					</hp:tc>`;
}

function renderChoiceTable({ choices, choiceNumerals, nextParagraphId, nextTableId }) {
	const chunks = chunkChoices(choices, 5);
	const tableId = nextTableId();
	const rows = chunks.map((chunk, rowIndex) => {
		const cells = chunk.map((choice, index) => {
			const globalIndex = rowIndex * 5 + index;
			const numeral = generateChoiceNumeral(choice, globalIndex, choiceNumerals);
			return renderChoiceTableCells({ numeral, text: choice, nextParagraphId, columnOffset: index * 2, rowIndex });
		});

		const missing = 5 - chunk.length;
		if (missing > 0) {
			for (let i = 0; i < missing; i += 1) {
				const paragraphA = nextParagraphId();
				const paragraphB = nextParagraphId();
				const columnOffset = (chunk.length + i) * 2;
				cells.push(`          <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="22">
						<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
							<hp:p id="${paragraphA}" paraPrIDRef="44" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
								<hp:run charPrIDRef="53" />
${renderLinesegArray({ width: '1440', height: '1150', baseline: '978' })}
							</hp:p>
						</hp:subList>
								<hp:cellAddr colAddr="${columnOffset}" rowAddr="${rowIndex}" />
						<hp:cellSpan colSpan="1" rowSpan="1" />
						<hp:cellSz width="1379" height="1431" />
						<hp:cellMargin left="0" right="0" top="0" bottom="0" />
					</hp:tc>
					<hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="22">
						<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
							<hp:p id="${paragraphB}" paraPrIDRef="44" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
								<hp:run charPrIDRef="53" />
${renderLinesegArray({ width: '4560', height: '1150', baseline: '978' })}
							</hp:p>
						</hp:subList>
							<hp:cellAddr colAddr="${columnOffset + 1}" rowAddr="${rowIndex}" />
						<hp:cellSpan colSpan="1" rowSpan="1" />
						<hp:cellSz width="4744" height="1431" />
						<hp:cellMargin left="184" right="0" top="0" bottom="0" />
					</hp:tc>`);
			}
		}

		return `        <hp:tr>
${cells.join('\n')}
				</hp:tr>`;
	}).join('\n');

	const paragraphId = nextParagraphId();
	const totalRows = chunks.length || 1;

	return `  <hp:p id="${paragraphId}" paraPrIDRef="60" styleIDRef="10" pageBreak="0" columnBreak="0" merged="0">
		<hp:run charPrIDRef="0">
			<hp:tbl id="${tableId}" zOrder="20" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="1" rowCnt="${totalRows}" colCnt="10" cellSpacing="0" borderFillIDRef="5" noAdjust="0">
				<hp:sz width="30615" widthRelTo="ABSOLUTE" protect="1" />
				<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0" />
				<hp:outMargin left="0" right="0" top="0" bottom="0" />
				<hp:inMargin left="0" right="0" top="0" bottom="0" />
${rows}
			</hp:tbl>
		</hp:run>
		<hp:run charPrIDRef="64">
			<hp:t />
		</hp:run>
${renderLinesegArray({ width: '30558', baseline: '1216' })}
	</hp:p>`;
}

function renderChoiceParagraphs({ choices, choiceNumerals, nextParagraphId }) {
	return choices.map((choice, index) => {
		const paragraphId = nextParagraphId();
		const numeral = generateChoiceNumeral(choice, index, choiceNumerals);
		return `  <hp:p id="${paragraphId}" paraPrIDRef="6" styleIDRef="10" pageBreak="0" columnBreak="0" merged="0">
		<hp:run charPrIDRef="2">
			<hp:t>${escapeXml(`${numeral} ${choice}`)}</hp:t>
		</hp:run>
${renderLinesegArray({ width: '31116', height: '1150', baseline: '978' })}
	</hp:p>`;
	}).join('\n');
}

function renderChoiceBlock({ question, options, nextParagraphId, nextTableId }) {
	const layout = question.choiceLayout || options.choiceLayout || (question.statementTitle ? 'table' : 'paragraph');

	if (layout === 'table') {
		return renderChoiceTable({ choices: question.choices, choiceNumerals: options.choiceNumerals, nextParagraphId, nextTableId });
	}

	return renderChoiceParagraphs({ choices: question.choices, choiceNumerals: options.choiceNumerals, nextParagraphId });
}

function renderSpacerParagraph(paragraphId) {
	return `  <hp:p id="${paragraphId}" paraPrIDRef="6" styleIDRef="10" pageBreak="0" columnBreak="0" merged="0">
		<hp:run charPrIDRef="2">
			<hp:t />
		</hp:run>
${renderLinesegArray({ width: '30558', height: '1150', baseline: '978' })}
	</hp:p>`;
}

function renderQuestionBlock({ question, questionIndex, options, nextParagraphId, nextTableId }) {
	const promptParagraph = renderPromptParagraph({
		paragraphId: nextParagraphId(),
		prompt: question.prompt,
		isFirstQuestion: questionIndex === 0,
		nextTableId
	});

	const contextEntries = normalizeContextEntries(question.contextEntries);
	const contextTable = renderContextTable({ contextEntries, nextParagraphId, nextTableId });

	const statementsTable = renderStatementsTable({
		statements: toArray(question.statements),
		statementTitle: question.statementTitle,
		nextParagraphId,
		nextTableId
	});

	const choicesBlock = renderChoiceBlock({ question, options, nextParagraphId, nextTableId });

	const spacerParagraphs = [];
	for (let i = 0; i < options.spacersPerQuestion; i += 1) {
		spacerParagraphs.push(renderSpacerParagraph(nextParagraphId()));
	}

	return [promptParagraph, contextTable, statementsTable, choicesBlock, spacerParagraphs.join('\n')]
		.filter(Boolean)
		.join('\n');
}

function renderSectionRoot(body) {
	return `${XML_DECLARATION}
<hs:sec xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core">
${body}
</hs:sec>`;
}

function renderAnswerHeader({ paragraphId, tableId }) {
	return `  <hp:p id="${paragraphId}" paraPrIDRef="70" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
		<hp:run charPrIDRef="53">
			<hp:rect id="${tableId}" zOrder="24" numberingType="PICTURE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" href="" groupLevel="0" instid="891785760" ratio="10">
				<hp:offset x="0" y="0" />
				<hp:orgSz width="8504" height="8504" />
				<hp:curSz width="66614" height="7200" />
				<hp:flip horizontal="0" vertical="0" />
				<hp:rotationInfo angle="0" centerX="33307" centerY="3600" rotateimage="1" />
				<hp:renderingInfo>
					<hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0" />
					<hc:scaMatrix e1="7.833255" e2="0" e3="0" e4="0" e5="0.84666" e6="0" />
					<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0" />
				</hp:renderingInfo>
				<hp:lineShape color="#000000" width="172" style="SOLID" endCap="FLAT" headStyle="NORMAL" tailStyle="NORMAL" headfill="1" tailfill="1" headSz="SMALL_SMALL" tailSz="SMALL_SMALL" outlineStyle="NORMAL" alpha="0" />
				<hp:shadow type="NONE" color="#B2B2B2" offsetX="0" offsetY="0" alpha="0" />
				<hp:drawText lastWidth="66614" name="" editable="0">
					<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
						<hp:p id="${paragraphId + 1}" paraPrIDRef="46" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
							<hp:run charPrIDRef="53">
								<hp:t>성공적인 직업생활 정답</hp:t>
							</hp:run>
						</hp:p>
					</hp:subList>
					<hp:textMargin left="283" right="283" top="283" bottom="283" />
				</hp:drawText>
				<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="0" allowOverlap="1" holdAnchorAndSO="0" vertRelTo="PAPER" horzRelTo="PAPER" vertAlign="TOP" horzAlign="CENTER" vertOffset="9895" horzOffset="0" />
				<hp:outMargin left="0" right="0" top="0" bottom="1417" />
			</hp:rect>
		</hp:run>
		<hp:run charPrIDRef="53">
			<hp:rect id="${tableId + 1}" zOrder="27" numberingType="PICTURE" textWrap="IN_FRONT_OF_TEXT" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" href="" groupLevel="0" instid="891785765" ratio="50">
				<hp:offset x="0" y="0" />
				<hp:orgSz width="8504" height="8504" />
				<hp:curSz width="14435" height="2331" />
				<hp:flip horizontal="0" vertical="0" />
				<hp:rotationInfo angle="0" centerX="7217" centerY="1165" rotateimage="1" />
				<hp:renderingInfo>
					<hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0" />
					<hc:scaMatrix e1="1.697437" e2="0" e3="0" e4="0" e5="0.274106" e6="0" />
					<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0" />
				</hp:renderingInfo>
				<hp:lineShape color="#000000" width="99" style="SOLID" endCap="FLAT" headStyle="NORMAL" tailStyle="NORMAL" headfill="1" tailfill="1" headSz="SMALL_SMALL" tailSz="SMALL_SMALL" outlineStyle="NORMAL" alpha="0" />
				<hp:shadow type="NONE" color="#B2B2B2" offsetX="0" offsetY="0" alpha="0" />
				<hp:drawText lastWidth="14435" name="" editable="0">
					<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
						<hp:p id="${paragraphId + 2}" paraPrIDRef="46" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
							<hp:run charPrIDRef="53">
								<hp:t>직업탐구</hp:t>
							</hp:run>
						</hp:p>
					</hp:subList>
					<hp:textMargin left="0" right="0" top="0" bottom="0" />
				</hp:drawText>
				<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="0" allowOverlap="1" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0" />
				<hp:outMargin left="0" right="0" top="0" bottom="850" />
			</hp:rect>
		</hp:run>
	</hp:p>`;
}

function renderAnswerEntry({ question, index, options, nextParagraphId, nextTableId }) {
	const rowParagraph = nextParagraphId();
	const tableId = nextTableId();
	const numeralParagraph = nextParagraphId();
	const textParagraph = nextParagraphId();
	const explanationParagraph = nextParagraphId();
	const answerLine = `${index + 1}. 정답: ${question.answer ?? ''}`.trim();
	const explanationLine = question.explanation ? `해설: ${question.explanation}` : '';

	const tableRow = `        <hp:tr>
					<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="22">
						<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
							<hp:p id="${numeralParagraph}" paraPrIDRef="44" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
								<hp:run charPrIDRef="53">
									<hp:t>${escapeXml(String(index + 1))}</hp:t>
								</hp:run>
${renderLinesegArray({ width: '1440', height: '1150', baseline: '978' })}
							</hp:p>
						</hp:subList>
						<hp:cellAddr colAddr="0" rowAddr="0" />
						<hp:cellSpan colSpan="1" rowSpan="1" />
						<hp:cellSz width="1379" height="1431" />
						<hp:cellMargin left="0" right="0" top="0" bottom="0" />
					</hp:tc>
					<hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="22">
						<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">
							<hp:p id="${textParagraph}" paraPrIDRef="44" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
								<hp:run charPrIDRef="53">
									<hp:t>${escapeXml(answerLine)}</hp:t>
								</hp:run>
${renderLinesegArray({ width: '4560', height: '1150', baseline: '978' })}
							</hp:p>
						</hp:subList>
						<hp:cellAddr colAddr="1" rowAddr="0" />
						<hp:cellSpan colSpan="9" rowSpan="1" />
						<hp:cellSz width="29236" height="1431" />
						<hp:cellMargin left="184" right="0" top="0" bottom="0" />
					</hp:tc>
				</hp:tr>`;

	const table = `  <hp:p id="${rowParagraph}" paraPrIDRef="60" styleIDRef="10" pageBreak="0" columnBreak="0" merged="0">
		<hp:run charPrIDRef="0">
			<hp:tbl id="${tableId}" zOrder="29" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="1" rowCnt="1" colCnt="10" cellSpacing="0" borderFillIDRef="3" noAdjust="1">
				<hp:sz width="30618" widthRelTo="ABSOLUTE" height="1431" heightRelTo="ABSOLUTE" protect="0" />
				<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0" />
				<hp:outMargin left="140" right="140" top="140" bottom="140" />
				<hp:inMargin left="140" right="140" top="140" bottom="140" />
${tableRow}
			</hp:tbl>
			<hp:t />
		</hp:run>
${renderLinesegArray({ width: '30700', height: '1431', baseline: '1216' })}
	</hp:p>`;

	const explanation = explanationLine
		? `  <hp:p id="${explanationParagraph}" paraPrIDRef="69" styleIDRef="48" pageBreak="0" columnBreak="0" merged="0">
		<hp:run charPrIDRef="77">
			<hp:t>${escapeXml(explanationLine)}</hp:t>
		</hp:run>
${renderLinesegArray({ width: '31116', height: '1100', baseline: '935' })}
	</hp:p>`
		: '';

	return `${table}
${explanation}`;
}

function renderAnswerSpacer(paragraphId) {
	return `  <hp:p id="${paragraphId}" paraPrIDRef="6" styleIDRef="10" pageBreak="0" columnBreak="0" merged="0">
		<hp:run charPrIDRef="2">
			<hp:t />
		</hp:run>
${renderLinesegArray({ width: '30558', height: '1150', baseline: '978' })}
	</hp:p>`;
}

function renderAnswerSection({ questions, options }) {
	const answerable = questions.filter((question) => question.answer || question.explanation);
	if (answerable.length === 0) {
		return null;
	}

	const nextParagraphId = createIdCounter(options.baseAnswerParagraphId);
	const nextTableId = createIdCounter(options.baseAnswerTableId);

	const header = renderAnswerHeader({ paragraphId: nextParagraphId(), tableId: nextTableId() });

	const entries = answerable.map((question, index) => {
		const block = renderAnswerEntry({ question, index, options, nextParagraphId, nextTableId });

		const spacers = [];
		for (let i = 0; i < options.answerSpacersPerQuestion; i += 1) {
			spacers.push(renderAnswerSpacer(nextParagraphId()));
		}

		return `${block}
${spacers.join('\n')}`;
	}).join('\n');

	const body = `${header}
${entries}`;

	return renderSectionRoot(body);
}

export function generateSectionXml({ questions, options } = {}) {
	const validatedQuestions = ensureQuestions(questions);
	const normalizedOptions = coerceOptions(options);

	const nextParagraphId = createIdCounter(normalizedOptions.baseParagraphId);
	const nextTableId = createIdCounter(normalizedOptions.baseTableId);

	const questionBlocks = validatedQuestions.map((question, index) => renderQuestionBlock({
		question,
		questionIndex: index,
		options: normalizedOptions,
		nextParagraphId,
		nextTableId
	}));

	const sectionBody = questionBlocks.join('\n');
	const sectionXml = renderSectionRoot(sectionBody);
	const answerSectionXml = renderAnswerSection({ questions: validatedQuestions, options: normalizedOptions });

	return { sectionXml, answerSectionXml };
}

