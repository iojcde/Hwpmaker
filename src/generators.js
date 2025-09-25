import minifyXml from 'minify-xml';
import { SECTION_OPEN, SECTION_CLOSE, DEFAULT_MINIFY_OPTIONS, CHOICE_LAYOUTS } from './constants.js';
import { createIdGenerator } from './utils.js';
import { buildParagraph, buildSectionOpeningParagraph, buildPromptParagraph, buildSpacerParagraph, buildChoiceParagraph } from './builders/paragraphs.js';
import { buildContextBlocks } from './builders/context.js';
import { buildStatementTable } from './builders/statements.js';
import { buildChoiceTable, formatChoiceText, normalizeChoice } from './builders/choices.js';

function normalizeQuestion(question, index) {
  if (!question || typeof question !== 'object') {
    throw new Error(`Question at index ${index} must be an object.`);
  }

  if (!question.prompt || typeof question.prompt !== 'string' || question.prompt.trim() === '') {
    throw new Error(`Question at index ${index} is missing a prompt.`);
  }

  const normalizedQuestion = {
    prompt: question.prompt.trim(),
    contextEntries: Array.isArray(question.contextEntries) ? question.contextEntries : [],
    statementTitle: question.statementTitle ?? null,
    statements: Array.isArray(question.statements) ? question.statements : [],
    choices: Array.isArray(question.choices) ? question.choices : [],
    answer: question.answer ?? null,
    explanation: question.explanation ?? null
  };

  // Determine choice layout
  if (question.choiceLayout) {
    normalizedQuestion.choiceLayout = question.choiceLayout;
  } else if (normalizedQuestion.statementTitle || normalizedQuestion.statements.length > 0) {
    normalizedQuestion.choiceLayout = CHOICE_LAYOUTS.TABLE;
  } else {
    normalizedQuestion.choiceLayout = CHOICE_LAYOUTS.PARAGRAPH;
  }

  return normalizedQuestion;
}

function generateSectionXml({ questions, options = {} }) {
  if (!Array.isArray(questions) || questions.length === 0) {
    throw new Error('"questions" must be a non-empty array.');
  }

  const { minifyOutput = true, ...generatorOptions } = options;

  const paragraphIdFactory = createIdGenerator(generatorOptions.baseParagraphId);
  const tableIdFactory = createIdGenerator(generatorOptions.baseTableId ?? 1900000000);

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
      const contextBlocks = buildContextBlocks({
        entries: question.contextEntries,
        paragraphIdFactory,
        tableIdFactory
      });
      blocks.push(...contextBlocks);
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
          options: generatorOptions,
          paragraphIdFactory
        }));
      } else {
        question.choices.forEach((choice, idx) => {
          const paragraphId = paragraphIdFactory();
          const choiceText = normalizeChoice(choice);
          const choiceContent = formatChoiceText(choiceText, idx, generatorOptions);
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

    const spacerCount = generatorOptions.spacersPerQuestion ?? 1;
    for (let i = 0; i < spacerCount; i += 1) {
      blocks.push(buildSpacerParagraph(paragraphIdFactory()));
    }
  });

  blocks.push(SECTION_CLOSE);
  const rawXml = `${blocks.join('\n')}`;
  return minifyOutput ? minifyXml(rawXml, DEFAULT_MINIFY_OPTIONS) : rawXml;
}

export {
  normalizeQuestion,
  generateSectionXml
};