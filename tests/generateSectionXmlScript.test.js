import test from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const SCRIPT_PATH = path.resolve(__dirname, '../scripts/generateSectionXml.js');

test('generateSectionXml script prints XML to stdout when requested', () => {
  const result = spawnSync('node', [SCRIPT_PATH, '--sample', '--stdout'], {
    encoding: 'utf8'
  });

  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /<hs:sec/);
  assert.match(result.stdout, /샘플 문제/);
  assert.match(result.stdout, /<hp:tbl[^>]+borderFillIDRef="7"/);
  assert.match(result.stdout, /<hp:tbl[^>]+borderFillIDRef="5"/);
  assert.match(result.stdout, /<hp:tc[^>]+borderFillIDRef="22"/);
  assert.match(result.stdout, /paraPrIDRef="55"/);
  assert.match(result.stdout, /colCnt="10"/);
});

test('generateSectionXml script writes XML file when output is provided', () => {
  const tempDir = mkdtempSync(path.join(os.tmpdir(), 'hwp-script-'));
  const outputPath = path.join(tempDir, 'section.xml');

  const result = spawnSync('node', [SCRIPT_PATH, '--sample', '--output', outputPath, '--minify'], {
    encoding: 'utf8'
  });

  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /Generated \d+ question\(s\)/);

  const xml = readFileSync(outputPath, 'utf8');
  assert.match(xml, /^<\?xml/);
  assert.match(xml, /<hs:sec/);
  assert.match(xml, /<hp:tbl[^>]+borderFillIDRef="7"/);
  assert.match(xml, /<hp:tbl[^>]+borderFillIDRef="5"/);
  assert.match(xml, /<hp:tc[^>]+borderFillIDRef="22"/);
  assert.match(xml, /paraPrIDRef="55"/);
  assert.match(xml, /colCnt="10"/);

  rmSync(tempDir, { recursive: true, force: true });
});
