import test from 'node:test';
import assert from 'node:assert/strict';

import { generateSectionXml, DEFAULT_CHOICE_NUMERALS } from '../src/hwpGenerator.js';

test('generateSectionXml throws for invalid input', () => {
  assert.throws(() => generateSectionXml({ questions: [] }), {
    name: 'Error',
    message: '"questions" must be a non-empty array.'
  });

  assert.throws(() => generateSectionXml({ questions: [{}] }), {
    name: 'Error',
    message: /missing a prompt/
  });
});

test('generateSectionXml creates XML for a single question', () => {
  const xml = generateSectionXml({
    questions: [
      {
        prompt: '다음 사례에 나타난 직업 가치관에 대한 설명으로 옳은 것은?',
        contextEntries: [
          { label: 'A 씨', text: '가업을 계승하여 자랑스럽고 복리 후생을 중시한다.' },
          { label: 'B 씨', text: '새로운 경험을 즐기며 상품 기획 전문가가 되었다.' }
        ],
        statementTitle: '&lt;보기&gt;',
        statements: [
          'A 씨는 귀속주의적 직업관을 가진다.',
          'B 씨는 변화 지향 가치를 중시한다.',
          'A 씨는 내재적 가치를, B 씨는 외재적 가치를 중시한다.'
        ],
        choices: [
          'ㄱ',
          'ㄴ',
          'ㄷ',
          'ㄱ, ㄴ',
          'ㄱ, ㄴ, ㄷ'
        ],
        answer: '④',
        explanation: 'A 씨는 귀속주의적 가치와 내재적 가치를 동시에 중시한다.'
      }
    ]
  });

  assert.ok(xml.startsWith('<?xml'), 'includes xml declaration');
  assert.ok(xml.includes('<hs:sec'), 'includes section root');
  assert.ok(xml.includes('다음 사례에 나타난 직업 가치관에 대한 설명으로 옳은 것은?'));
  assert.match(xml, /<hp:tbl[^>]+rowCnt="2"[^>]+borderFillIDRef="7"/);
  assert.match(xml, /<hp:tbl[^>]+rowCnt="3"[^>]+colCnt="3"[^>]+borderFillIDRef="22"/);
  assert.match(xml, /colCnt="10"/);
  assert.ok(xml.includes('<hp:cellSz width="30611"'));
  assert.ok(xml.includes('&lt;보기&gt;'));
  assert.ok(xml.includes('A 씨 : 가업을 계승하여 자랑스럽고 복리 후생을 중시한다.'));
  assert.ok(xml.includes('A 씨는 귀속주의적 직업관을 가진다.'));
  assert.ok(xml.includes('B 씨는 변화 지향 가치를 중시한다.'));
  assert.ok(xml.includes('A 씨는 내재적 가치를, B 씨는 외재적 가치를 중시한다.'));
  assert.ok(xml.includes('ㄱ, ㄴ, ㄷ'));
  assert.ok(xml.includes('정답: ④'));
  assert.ok(xml.includes('A 씨는 귀속주의적 가치와 내재적 가치를 동시에 중시한다.'));
});

test('generateSectionXml supports multiple questions and custom choice numerals', () => {
  const xml = generateSectionXml({
    questions: [
      {
        prompt: '첫 번째 문항',
        choices: ['선택 1', '선택 2']
      },
      {
        prompt: '두 번째 문항',
        choices: ['옵션 A', '옵션 B', '옵션 C']
      }
    ],
    options: {
      choiceNumerals: ['(1)', '(2)', '(3)', '(4)']
    }
  });

  assert.ok(xml.includes('첫 번째 문항'));
  assert.ok(xml.includes('두 번째 문항'));
  assert.ok(xml.includes('(1) 선택 1'));
  assert.ok(xml.includes('(3) 옵션 C'));
  DEFAULT_CHOICE_NUMERALS.slice(0, 2).forEach((numeral) => {
    assert.ok(!xml.includes(`${numeral} 선택 1`));
  });
});
