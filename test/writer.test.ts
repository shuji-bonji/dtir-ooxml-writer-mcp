/**
 * vitest spec — dtirToDocx の主要不変条件（reader 非依存・静的 DTIR フィクスチャ）。
 * LibreOffice 込みの受け入れは test/roundtrip.ts（npm run test:roundtrip）。
 */
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { fixtureDocxPath, readerDtirPath } from '@shuji-bonji/doc-translation-ir/fixtures';
import type { IRDocument } from '@shuji-bonji/doc-translation-ir';
import { dtirToDocx } from '../src/writer.js';

/** 同梱の reader 出力 DTIR を毎回フレッシュに読む（テスト間で破壊的更新するため）。 */
function loadDtir(): IRDocument {
  return JSON.parse(readFileSync(readerDtirPath, 'utf8')) as IRDocument;
}
const orig = readFileSync(fixtureDocxPath);

async function docXml(buf: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(buf);
  return (await zip.file('word/document.xml')?.async('string')) ?? '';
}

function fill(dtir: IRDocument, text: (s: string) => string): IRDocument {
  for (const s of dtir.segments) {
    if (s.translatable) {
      s.translation = {
        text: text(s.text.source),
        engine: 'pseudo',
        sourceLangUsed: s.language.value,
        targetLang: 'en-GB',
        at: new Date().toISOString(),
      };
    }
  }
  return dtir;
}

describe('dtirToDocx', () => {
  it('patches translatable text, preserves fields/numerics/structure', async () => {
    const dtir = fill(loadDtir(), (s) => `EN〔${s}〕`);
    const out = await dtirToDocx(dtir, orig, { onMissingTranslation: 'error' });
    const xml = await docXml(out);
    expect(xml).toContain('EN〔Jaarverslag 2025〕');
    expect(xml).toContain('Resultaten'); // フィールドキャッシュは不可触
    expect(xml).toContain('1.250.000'); // 数値は不可触
    expect(xml).toContain('<w:sectPr'); // 構造保持
  });

  it('keeps original text when translation missing (keep mode)', async () => {
    const out = await dtirToDocx(loadDtir(), orig, { onMissingTranslation: 'keep' });
    const xml = await docXml(out);
    expect(xml).toContain('Jaarverslag 2025'); // 未翻訳なら原文維持
  });

  it('collapses run-split paragraphs (empty w:t generated)', async () => {
    const dtir = fill(loadDtir(), () => 'X');
    const out = await dtirToDocx(dtir, orig, { onMissingTranslation: 'keep' });
    const xml = await docXml(out);
    // 仏語の3ランが先頭1本に集約され、残り2ランの w:t が空になる
    const empty = (xml.match(/<w:t\b[^>]*\/>|<w:t\b[^>]*><\/w:t>/g) ?? []).length;
    expect(empty).toBeGreaterThanOrEqual(2);
  });

  it('throws when a translatable segment lacks translation in error mode', async () => {
    await expect(dtirToDocx(loadDtir(), orig, { onMissingTranslation: 'error' })).rejects.toThrow();
  });
});
