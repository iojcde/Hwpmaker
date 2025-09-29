# Hwpmaker XML Generator

Generate Hangul Word Processor (HWPML) section XML from structured question data. The library exposes a single entry point, `generateSectionXml`, that composes layout primitives similar to the sample `section0.xml` provided in this repository while also producing a companion answer section (`section1.xml`) when answers or explanations are present. It lets you author question banks programmatically while retaining consistent formatting in the resulting `.hwpx` documents.

## Installation

Install dependencies with Bun if it's available on your system; otherwise fall back to pnpm (bundled with the dev container). The project already includes `minify-xml` for optional post-processing.

```bash
bun install
```

```bash
pnpm install
```

## Usage

### Library API

```js
import { generateSectionXml } from './src/hwpGenerator.js';
import { writeFileSync } from 'node:fs';

const { sectionXml, answerSectionXml } = generateSectionXml({
  questions: [
    {
      prompt: '다음 사례에 나타난 직업 가치관에 대한 설명으로 옳은 것은?',
      contextEntries: [
        { label: 'A 씨', text: '가업을 계승하며 복리 후생을 중시한다.' },
        { label: 'B 씨', text: '새로운 경험을 즐기며 상품 기획 전문가가 되었다.' }
      ],
      statementTitle: '&lt;보 기&gt;',
      statements: [
        'A 씨는 귀속주의적 직업관을 가진다.',
        'B 씨는 변화 지향 가치를 중시한다.',
        'A 씨는 내재적 가치를, B 씨는 외재적 가치를 중시한다.'
      ],
      choices: ['ㄱ', 'ㄴ', 'ㄷ', 'ㄱ, ㄴ', 'ㄱ, ㄴ, ㄷ'],
      answer: '④',
      explanation: '귀속주의적 가치와 내재적 가치를 함께 강조한다.'
    }
  ],
  options: {
    choiceNumerals: ['①', '②', '③', '④', '⑤'],
    spacersPerQuestion: 2
  }
});

writeFileSync('section0.xml', sectionXml, 'utf-8');
if (answerSectionXml) {
  writeFileSync('section1.xml', answerSectionXml, 'utf-8');
}
```

### CLI generator

Generate a section XML file directly from the command line. The CLI accepts a JSON payload that matches the structure described above or you can use the built-in sample payload for a quick smoke test.

```bash
node scripts/generateSectionXml.js --sample --output file/Contents/section0.xml
```

The CLI writes the main question section to the requested output path and automatically emits the answer section alongside it (e.g. `section1.xml`).

Key options:

- `--input <path>`: JSON file containing a `questions` array and optional `options` object.
- `--options <json>`: Inline JSON string merged into generator options (e.g. `{"spacersPerQuestion":0}`).
- `--minify`: Compress the output using `minify-xml`.
- `--stdout`: Print to standard output instead of writing a file.
- `--answers-output <path>`: Explicit path for the generated answer section (defaults to `section1.xml` next to the main output).

### Question shape

Each question supports the following fields:

| Field | Type | Description |
| --- | --- | --- |
| `prompt` | `string` | Required. Main question text. |
| `contextEntries` | `Array<{label?: string, text: string}>` | Optional. Each entry becomes a row inside the boxed table shown in the reference layout (e.g., 사례 A, B). |
| `statementTitle` | `string` | Optional label rendered in the narrow middle column of the 보기 table (e.g. `&lt;보기&gt;`). The generator automatically converts `&lt;`/`&gt;` entities into visible brackets. |
| `statements` | `string[]` | Optional bullet-style statements occupying the merged bottom row of the 보기 table, replicating the boxed layout shown in the reference file. |
| `choices` | `string[]` | Optional answers. Rendered either as numbered paragraphs or as the paired table shown above depending on `choiceLayout`. Provide plain text (e.g. `ㄱ, ㄴ`) and the generator adds numbering when needed. |
| `choiceLayout` | `'paragraph' \| 'table'` | Optional. Defaults to `'table'` when `statementTitle` is present (보기 style) and `'paragraph'` otherwise. Set explicitly to control the answer layout. |
| `answer` | `string` | Optional answer footer (rendered as `정답: ...`). Appears in the separate answer section (`section1.xml`). |
| `explanation` | `string` | Optional freestyle paragraph after the answer, also emitted in the answer section. |

### Options

| Option | Default | Description |
| --- | --- | --- |
| `choiceNumerals` | Korean circled numerals ①–⑩ | Prefix sequence for choices. Extra entries fall back to numeric counters. |
| `spacersPerQuestion` | `1` | Number of blank paragraphs appended after each question block. |
| `baseParagraphId` | `2147483648` | Starting numeric seed for generated paragraph IDs. |
| `baseTableId` | `1900000000` | Starting numeric seed for internal table IDs inserted into the section header. |
| `choiceLayout` | (per-question) | Override default layout for all choices (e.g. force `'paragraph'` in options or question payload). |
| `answerHeading` | `'정답 및 해설'` | Heading text placed at the top of the generated answer section. |
| `answerSpacersPerQuestion` | matches `spacersPerQuestion` | Number of blank paragraphs appended after each answer/explanation pair. |
| `baseAnswerParagraphId` | inherits `baseParagraphId` | Starting seed for paragraph IDs used in the answer section. |
| `baseAnswerTableId` | inherits `baseTableId` | Starting seed for table IDs used when laying out the answer section header. |

## Tests

Run the automated tests with Node’s built-in runner:

```bash
pnpm test
```

## Next steps

- Hook the generator into a CLI or web UI for authoring question banks.
- Post-process the output with `minify-xml` for smaller package sizes if you embed XML in distributions.
- Extend the generator with richer table layouts when you need multi-column choices or matrix questions.
