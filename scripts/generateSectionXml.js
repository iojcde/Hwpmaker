#!/usr/bin/env node

import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

import minifyXml from 'minify-xml';

import { generateSectionXml, DEFAULT_MINIFY_OPTIONS } from '../src/hwpGenerator.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const DEFAULT_OUTPUT = path.resolve(__dirname, '..', 'file', 'Contents', 'generated-section.xml');
const SAMPLE_PAYLOAD = {
  questions: [
    {
      prompt: '샘플 문제: 다음 중 XML 생성 스크립트의 장점을 모두 고르시오.',
      contextEntries: [
        {
          label: '',
          text: '스크립트는 JSON 파일을 읽어 HWP XML을 생성합니다.'
        },
        {
          label: '',
          table: {
            headers: ['제품명', '카테고리', '가격', '재고'],
            rows: [
              ['MacBook Pro', '노트북', '2,500,000', '15'],
              ['iPhone 15', '스마트폰', '1,200,000', '32'],
              ['AirPods Pro', '이어폰', '350,000', '48'],
              ['iPad Air', '태블릿', '800,000', '22']
            ]
          }
        }
      ],
      statementTitle: '<보 기>',
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
    },
    {
      prompt: '하하하 호호호',
      contextEntries: [
        {
          label: '',
          text: '여기는 description'
        },
        {
          label: '',
          table: {
            headers: ['연도', '매출'],
            rows: [
              ['2023', '120'],
              ['2024', '180']
            ]
          }
        }
      ],
      statementTitle: '<보 기>',
      statements: [
        '하하하.',
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
      explanation: '스크립트는 반복 작업을 자동화하고 확장하기 안 쉽습니다.'
    },
    {
      prompt: '다음 중 클라우드 컴퓨팅의 장점으로 옳은 것만 고르시오.',
      contextEntries: [
        {
          label: '',
          text: '클라우드 서비스는 확장성과 유연성을 제공합니다.'
        },
        {
          label: '',
          table: {
            headers: ['서비스', '특징'],
            rows: [
              ['IaaS', '가상 서버 제공'],
              ['PaaS', '개발 환경 제공'],
              ['SaaS', '애플리케이션 제공']
            ]
          }
        }
      ],
      statementTitle: '<보 기>',
      statements: [
        '사용자가 직접 서버를 구매할 필요가 없다.',
        '필요에 따라 자원을 쉽게 늘릴 수 있다.',
        '모든 클라우드 서비스는 무료이다.'
      ],
      choiceLayout: 'table',
      choices: ['ㄱ', 'ㄴ', 'ㄷ', 'ㄱ, ㄴ', 'ㄱ, ㄴ, ㄷ'],
      answer: '④',
      explanation: '클라우드의 주요 장점은 비용 절감과 확장성 제공입니다.'
    },
    {
      prompt: '데이터베이스 인덱스의 특징으로 옳은 것은?',
      contextEntries: [
        {
          label: '',
          text: '인덱스는 데이터 검색 성능을 향상시키는 자료 구조입니다.'
        }
      ],
      statementTitle: '<보 기>',
      statements: [
        '검색 속도를 높일 수 있다.',
        '데이터 삽입 시 항상 성능이 향상된다.',
        '디스크 공간을 추가로 사용한다.'
      ],
      choiceLayout: 'table',
      choices: ['ㄱ', 'ㄴ', 'ㄷ', 'ㄱ, ㄷ', 'ㄱ, ㄴ, ㄷ'],
      answer: '④',
      explanation: '인덱스는 검색 성능을 높이지만 삽입 시 오버헤드가 발생합니다.'
    },
    {
      prompt: '다음 표를 참고하여 평균 점수를 고르시오.',
      contextEntries: [
        {
          label: '',
          table: {
            headers: ['학생', '국어', '수학'],
            rows: [
              ['A', '80', '90'],
              ['B', '70', '100'],
              ['C', '90', '95']
            ]
          }
        }
      ],
      statementTitle: '<보 기>',
      statements: [
        '국어 평균은 80점이다.',
        '수학 평균은 95점이다.',
        '전체 평균은 87.5점이다.'
      ],
      choiceLayout: 'table',
      choices: ['ㄱ', 'ㄴ', 'ㄷ', 'ㄱ, ㄴ', 'ㄱ, ㄷ'],
      answer: '⑤',
      explanation: '국어 평균 80, 수학 평균 95, 전체 평균은 87.5점입니다.'
    },
    {
      prompt: '네트워크 계층 구조에 대한 설명으로 옳은 것을 모두 고르시오.',
      contextEntries: [
        {
          label: '',
          text: 'OSI 7계층은 네트워크 동작을 이해하기 위한 개념적 모델입니다.'
        }
      ],
      statementTitle: '<보 기>',
      statements: [
        '전송 계층은 데이터의 신뢰성을 보장한다.',
        '네트워크 계층은 IP 주소를 사용한다.',
        '물리 계층은 애플리케이션 로직을 담당한다.'
      ],
      choiceLayout: 'table',
      choices: ['ㄱ', 'ㄴ', 'ㄷ', 'ㄱ, ㄴ', 'ㄱ, ㄴ, ㄷ'],
      answer: '④',
      explanation: '물리 계층은 하드웨어 신호를 다루고 애플리케이션 로직은 담당하지 않습니다.'
    },
    {
      prompt: '다음 중 머신러닝 모델의 특징으로 옳지 않은 것은?',
      contextEntries: [
        {
          label: '',
          text: '머신러닝은 데이터로부터 패턴을 학습합니다.'
        }
      ],
      statementTitle: '<보 기>',
      statements: [
        '지도 학습은 레이블이 있는 데이터를 사용한다.',
        '비지도 학습은 레이블이 없는 데이터를 사용한다.',
        '모든 모델은 항상 100% 정확도를 보장한다.'
      ],
      choiceLayout: 'table',
      choices: ['ㄱ', 'ㄴ', 'ㄷ', 'ㄱ, ㄴ', 'ㄱ, ㄴ, ㄷ'],
      answer: '②',
      explanation: '모든 모델이 항상 정확한 것은 아닙니다.'
    },
    {
      prompt: '다음 중 Git의 장점으로 옳은 것을 모두 고르시오.',
      contextEntries: [
        {
          label: '',
          text: 'Git은 분산 버전 관리 시스템입니다.'
        }
      ],
      statementTitle: '<보 기>',
      statements: [
        '여러 개발자가 동시에 작업할 수 있다.',
        '모든 변경 이력은 추적 가능하다.',
        '중앙 서버가 없으면 사용할 수 없다.'
      ],
      choiceLayout: 'table',
      choices: ['ㄱ', 'ㄴ', 'ㄷ', 'ㄱ, ㄴ', 'ㄱ, ㄴ, ㄷ'],
      answer: '④',
      explanation: 'Git은 분산형이므로 중앙 서버 없이도 사용 가능합니다.'
    },
    {
      prompt: '다음 중 HTML5의 새로운 기능에 해당하는 것을 고르시오.',
      contextEntries: [
        {
          label: '',
          text: 'HTML5는 멀티미디어와 구조적 요소를 강화했습니다.'
        },
        {
          label: '',
          table: {
            headers: ['태그', '용도'],
            rows: [
              ['<canvas>', '그래픽 그리기'],
              ['<video>', '비디오 재생'],
              ['<article>', '문서 구획']
            ]
          }
        }
      ],
      statementTitle: '<보 기>',
      statements: [
        '<canvas>는 그래픽을 그릴 수 있다.',
        '<video>는 비디오를 재생한다.',
        '<font>는 HTML5에서 새롭게 추가되었다.'
      ],
      choiceLayout: 'table',
      choices: ['ㄱ', 'ㄴ', 'ㄷ', 'ㄱ, ㄴ', 'ㄱ, ㄴ, ㄷ'],
      answer: '④',
      explanation: '<font> 태그는 HTML5에서 폐지되었습니다.'
    },
    {
      prompt: '데이터 시각화 도구에 대한 설명으로 옳은 것을 모두 고르시오.',
      contextEntries: [
        {
          label: '',
          text: '데이터 시각화는 복잡한 데이터를 쉽게 이해할 수 있도록 합니다.'
        },
        {
          label: '',
          table: {
            headers: ['도구', '특징'],
            rows: [
              ['Tableau', '시각적 대시보드'],
              ['Matplotlib', '파이썬 라이브러리'],
              ['Excel', '간단한 차트 생성']
            ]
          }
        }
      ],
      statementTitle: '<보 기>',
      statements: [
        'Tableau는 대화형 시각화에 강하다.',
        'Matplotlib은 주로 파이썬에서 사용된다.',
        'Excel은 차트 기능이 없다.'
      ],
      choiceLayout: 'table',
      choices: ['ㄱ', 'ㄴ', 'ㄷ', 'ㄱ, ㄴ', 'ㄱ, ㄴ, ㄷ'],
      answer: '④',
      explanation: 'Excel에도 기본 차트 기능이 있습니다.'
    }
  ]
};

function printHelp() {
  const relativeDefaultOutput = path.relative(process.cwd(), DEFAULT_OUTPUT);
  const usage = `Usage: generateSectionXml [options]\n\n` +
    'Options:\n' +
    '  -i, --input <path>       JSON file containing "questions" data\n' +
    '  -o, --output <path>      Destination XML file (defaults to ' + relativeDefaultOutput + ')\n' +
    '      --stdout             Print XML to stdout instead of writing a file\n' +
    '      --answers-output <path>  Destination XML file for generated answers section\n' +
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
    generatorOptions: {},
    answersOutput: null
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
      case '--answers-output':
        if (i + 1 >= argv.length) {
          throw new Error(`${arg} requires a value.`);
        }
        parsed.answersOutput = path.resolve(process.cwd(), argv[i + 1]);
        i += 1;
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

function deriveAnswerOutputPath(questionOutputPath) {
  const directory = path.dirname(questionOutputPath);
  const basename = path.basename(questionOutputPath);

  if (/section0\.xml$/i.test(basename)) {
    return path.join(directory, basename.replace(/section0\.xml$/i, 'section1.xml'));
  }

  if (basename.toLowerCase().endsWith('.xml')) {
    const stem = basename.slice(0, -4);
    return path.join(directory, `${stem}.answers.xml`);
  }

  return path.join(directory, `${basename}.answers.xml`);
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

  let generated;
  try {
    generated = generateSectionXml({
      questions: payload.questions,
      options: combinedOptions
    });
  } catch (error) {
    console.error(`[generateSectionXml] Failed to generate XML: ${error.message}`);
    process.exit(1);
  }

  const { sectionXml, answerSectionXml } = generated;

  const applyMinifyIfRequested = (xmlString) => {
    if (!xmlString) {
      return null;
    }

    if (!parsedArgs.minify) {
      return xmlString;
    }

    try {
      return minifyXml(xmlString, DEFAULT_MINIFY_OPTIONS);
    } catch (error) {
      console.error(`[generateSectionXml] Failed to minify XML: ${error.message}`);
      process.exit(1);
    }
  };

  const finalSectionXml = applyMinifyIfRequested(sectionXml);
  const finalAnswerXml = applyMinifyIfRequested(answerSectionXml);

  if (parsedArgs.stdout) {
    const stdoutXml = finalSectionXml ?? '';
    process.stdout.write(stdoutXml);
    if (!stdoutXml.endsWith('\n')) {
      process.stdout.write('\n');
    }
    return;
  }

  const outputPath = parsedArgs.output ?? DEFAULT_OUTPUT;
  try {
    await ensureDirectoryForFile(outputPath);
    await writeFile(outputPath, finalSectionXml ?? '', 'utf8');
  } catch (error) {
    console.error(`[generateSectionXml] Failed to write XML: ${error.message}`);
    process.exit(1);
  }

  let answersOutputPath = null;
  if (finalAnswerXml) {
    answersOutputPath = parsedArgs.answersOutput ?? deriveAnswerOutputPath(outputPath);
    try {
      await ensureDirectoryForFile(answersOutputPath);
      await writeFile(answersOutputPath, finalAnswerXml, 'utf8');
    } catch (error) {
      console.error(`[generateSectionXml] Failed to write answers XML: ${error.message}`);
      process.exit(1);
    }
  }

  const relativePath = path.relative(process.cwd(), outputPath);
  console.log(`Generated ${payload.questions.length} question(s) at ${relativePath}`);
  if (answersOutputPath) {
    const relativeAnswersPath = path.relative(process.cwd(), answersOutputPath);
    console.log(`Answer section written to ${relativeAnswersPath}`);
  }
  if (parsedArgs.minify) {
    console.log('Minification enabled (minify-xml options applied).');
  }
}

main().catch((error) => {
  console.error(`[generateSectionXml] Unexpected error: ${error.message}`);
  process.exit(1);
});
