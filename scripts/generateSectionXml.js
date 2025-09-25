#!/usr/bin/env node

import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

import minifyXml from 'minify-xml';

import { generateSectionXml } from '../src/hwpGenerator.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const DEFAULT_OUTPUT = path.resolve(__dirname, '..', 'file', 'Contents', 'generated-section.xml');
const SAMPLE_PAYLOAD = {
  questions: [
    {
      prompt: '샘플 문제: 다음 중 XML 생성 스크립트의 장점을 모두 고르시오.',
      contextEntries: [
        {
          label: '정보',
          text: '스크립트는 JSON 파일을 읽어 HWP XML을 생성합니다.'
        }
      ],
      statementTitle: '<보기>',
      statements: [
        '단일 명령으로 XML 파일을 만들 수 있다.',
        '여러 문제를 한 번에 생성하도록 확장할 수 있다.'
      ],
      choiceLayout: 'table',
      choices: [
        'ㄱ',
        'ㄴ',
        'ㄷ',
        'ㄱ, ㄴ',
        'ㄱ, ㄴ, ㄷ'
      ],
  answer: '⑤',
      explanation: '스크립트는 반복 작업을 자동화하고 확장하기 쉽습니다.'
    }
  ]
};
const DEFAULT_MINIFY_OPTIONS = {
  removeComments: true,
  removeWhitespaceBetweenTags: true,
  considerPreserveWhitespace: true,
  collapseWhitespaceInTags: true,
  collapseEmptyElements: true,
  trimWhitespaceFromTexts: false,
  collapseWhitespaceInTexts: false,
  collapseWhitespaceInProlog: true,
  collapseWhitespaceInDocType: true,
  removeSchemaLocationAttributes: false,
  removeUnnecessaryStandaloneDeclaration: true,
  removeUnusedNamespaces: false,
  removeUnusedDefaultNamespace: false,
  shortenNamespaces: false,
  ignoreCData: true
};

function printHelp() {
  const relativeDefaultOutput = path.relative(process.cwd(), DEFAULT_OUTPUT);
  const usage = `Usage: generateSectionXml [options]\n\n` +
    'Options:\n' +
    '  -i, --input <path>       JSON file containing "questions" data\n' +
    '  -o, --output <path>      Destination XML file (defaults to ' + relativeDefaultOutput + ')\n' +
    '      --stdout             Print XML to stdout instead of writing a file\n' +
    '      --options <json>     JSON string with generator options (e.g. {"spacersPerQuestion":0})\n' +
    '      --minify             Minify generated XML using minify-xml\n' +
    '      --sample             Use built-in sample question payload\n' +
    '      --base-paragraph <n> Set starting paragraph ID\n' +
    '      --base-table <n>     Set starting table ID\n' +
    '  -h, --help               Show this help message\n';

  process.stdout.write(usage);
}

function parseArgs(argv) {
  const parsed = {
    input: null,
    output: DEFAULT_OUTPUT,
    stdout: false,
    useSample: false,
    minify: false,
    generatorOptions: {}
  };

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    switch (arg) {
      case '-i':
      case '--input':
        if (i + 1 >= argv.length) {
          throw new Error(`${arg} requires a value.`);
        }
        parsed.input = argv[i + 1];
        i += 1;
        break;
      case '-o':
      case '--output':
        if (i + 1 >= argv.length) {
          throw new Error(`${arg} requires a value.`);
        }
        parsed.output = path.resolve(process.cwd(), argv[i + 1]);
        i += 1;
        break;
      case '--stdout':
        parsed.stdout = true;
        break;
      case '--options':
        if (i + 1 >= argv.length) {
          throw new Error(`${arg} requires a JSON value.`);
        }
        try {
          parsed.generatorOptions = {
            ...parsed.generatorOptions,
            ...JSON.parse(argv[i + 1])
          };
        } catch (error) {
          throw new Error(`Unable to parse value for --options: ${error.message}`);
        }
        i += 1;
        break;
      case '--minify':
        parsed.minify = true;
        break;
      case '--sample':
        parsed.useSample = true;
        break;
      case '--base-paragraph':
        if (i + 1 >= argv.length) {
          throw new Error(`${arg} requires a numeric value.`);
        }
        parsed.generatorOptions.baseParagraphId = Number.parseInt(argv[i + 1], 10);
        i += 1;
        break;
      case '--base-table':
        if (i + 1 >= argv.length) {
          throw new Error(`${arg} requires a numeric value.`);
        }
        parsed.generatorOptions.baseTableId = Number.parseInt(argv[i + 1], 10);
        i += 1;
        break;
      case '-h':
      case '--help':
        printHelp();
        process.exit(0);
        break;
      default:
        throw new Error(`Unknown argument: ${arg}`);
    }
  }

  return parsed;
}

async function ensureDirectoryForFile(targetPath) {
  const directory = path.dirname(targetPath);
  if (!existsSync(directory)) {
    await mkdir(directory, { recursive: true });
  }
}

async function loadPayload({ inputPath, useSample }) {
  if (useSample && inputPath) {
    throw new Error('Use either --sample or --input, not both.');
  }

  if (useSample || !inputPath) {
    return SAMPLE_PAYLOAD;
  }

  let contents;
  try {
    const resolvedPath = path.resolve(process.cwd(), inputPath);
    contents = await readFile(resolvedPath, 'utf8');
  } catch (error) {
    throw new Error(`Unable to read JSON input from ${inputPath}: ${error.message}`);
  }

  try {
    const parsed = JSON.parse(contents);
    if (Array.isArray(parsed)) {
      return { questions: parsed };
    }

    if (parsed && typeof parsed === 'object') {
      if (!parsed.questions && !Array.isArray(parsed)) {
        throw new Error('Input JSON must contain a "questions" array.');
      }
      const questions = Array.isArray(parsed.questions) ? parsed.questions : [];
      const options = parsed.options && typeof parsed.options === 'object' ? parsed.options : {};
      return { questions, options };
    }

    throw new Error('Input JSON must be an object or array.');
  } catch (error) {
    throw new Error(`Unable to parse JSON: ${error.message}`);
  }
}

async function main() {
  const argv = process.argv.slice(2);

  let parsedArgs;
  try {
    parsedArgs = parseArgs(argv);
  } catch (error) {
    console.error(`[generateSectionXml] ${error.message}`);
    printHelp();
    process.exit(1);
  }

  let payload;
  try {
    payload = await loadPayload({
      inputPath: parsedArgs.input,
      useSample: parsedArgs.useSample
    });
  } catch (error) {
    console.error(`[generateSectionXml] ${error.message}`);
    process.exit(1);
  }

  const combinedOptions = {
    ...(payload.options ?? {}),
    ...(parsedArgs.generatorOptions || {})
  };

  let xml;
  try {
    xml = generateSectionXml({
      questions: payload.questions,
      options: combinedOptions
    });
  } catch (error) {
    console.error(`[generateSectionXml] Failed to generate XML: ${error.message}`);
    process.exit(1);
  }

  let outputXml = xml;
  if (parsedArgs.minify) {
    try {
      outputXml = minifyXml(xml, DEFAULT_MINIFY_OPTIONS);
    } catch (error) {
      console.error(`[generateSectionXml] Failed to minify XML: ${error.message}`);
      process.exit(1);
    }
  }

  if (parsedArgs.stdout) {
    process.stdout.write(outputXml);
    if (!outputXml.endsWith('\n')) {
      process.stdout.write('\n');
    }
    return;
  }

  const outputPath = parsedArgs.output ?? DEFAULT_OUTPUT;
  try {
    await ensureDirectoryForFile(outputPath);
    await writeFile(outputPath, outputXml, 'utf8');
  } catch (error) {
    console.error(`[generateSectionXml] Failed to write XML: ${error.message}`);
    process.exit(1);
  }

  const relativePath = path.relative(process.cwd(), outputPath);
  console.log(`Generated ${payload.questions.length} question(s) at ${relativePath}`);
  if (parsedArgs.minify) {
    console.log('Minification enabled (minify-xml options applied).');
  }
}

main().catch((error) => {
  console.error(`[generateSectionXml] Unexpected error: ${error.message}`);
  process.exit(1);
});
