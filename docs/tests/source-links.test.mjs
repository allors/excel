import assert from 'node:assert/strict';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { rewriteLink } from '../integrations/source-links.mjs';

const source = fileURLToPath(new URL('../src/content/docs/internals/architecture.md', import.meta.url));

test('source links become site links at a project or domain root', () => {
  assert.equal(rewriteLink('../api.md#iworksheet', source, '/excel/'), '/excel/api/#iworksheet');
  assert.equal(rewriteLink('../index.md', source, '/'), '/');
  assert.equal(rewriteLink('#tests', source, '/excel/'), '#tests');
});

test('repository references resolve to the documented branch', () => {
  assert.equal(rewriteLink('../../../../../README.md#documentation', source, '/excel/'),
    'https://github.com/allors/excel/blob/main/README.md#documentation');
  assert.equal(rewriteLink('https://example.com/guide.md', source, '/excel/'), 'https://example.com/guide.md');
});

test('broken source links fail the build with a useful target', () => {
  assert.throws(() => rewriteLink('../missing.md', source, '/excel/'), /missing\.md/);
});

test('repository fragments are checked against the local Markdown headings', () => {
  assert.equal(rewriteLink('../../../../../AGENTS.md#documentation', source, '/excel/'),
    'https://github.com/allors/excel/blob/main/AGENTS.md#documentation');
  assert.throws(() => rewriteLink('../../../../../AGENTS.md#no-such-heading', source, '/excel/'),
    /no-such-heading.*AGENTS\.md/);
  const fixture = 'https://github.com/allors/excel/blob/main/docs/tests/fixtures/anchors.md';
  for (const id of ['added', 'added-1', 'use-rewritelink-in-astro', 'explicit-anchor']) {
    assert.equal(rewriteLink(`../../../../tests/fixtures/anchors.md#${id}`, source, '/excel/'), `${fixture}#${id}`);
  }
  assert.throws(() => rewriteLink('../../../../tests/fixtures/anchors.md#not-a-heading', source, '/excel/'), /not-a-heading/);
});

test('fragments on repository files that are not Markdown are left alone', () => {
  assert.equal(rewriteLink('../../../../../src/Allors.Excel.Interop/Worksheet.cs#L45', source, '/excel/'),
    'https://github.com/allors/excel/blob/main/src/Allors.Excel.Interop/Worksheet.cs#L45');
});
