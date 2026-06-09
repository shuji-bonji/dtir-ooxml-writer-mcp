/**
 * vitest spec — dtirToDocx の主要不変条件。
 * フル往復は test/roundtrip.ts（npm run test:roundtrip, LibreOffice 検証込み）。
 */
import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { docxToDtir } from '../../dtir-ooxml-reader-mcp/src/reader.js';
import { dtirToDocx } from '../src/writer.js';

const here = dirname(fileURLToPath(import.meta.url));
const fixture = resolve(here, '../../doc-translation-ir/fixtures/docx/mixed-nl-fr-de-tricky.docx');

async function docXml(buf: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(buf);
  return (await zip.file('word/document.xml')?.async('string')) ?? '';
}

describe('dtirToDocx', () => {
  it('patches translatable text, preserves fields/numerics/structure', async () => {
    const orig = readFileSync(fixture);
    const dtir = await docxToDtir(orig, { fileName: 'x.docx', targetLang: 'en-GB' });
    for (const s of dtir.segments) {
      if (s.translatable) {
        s.translation = {
          text: `EN〔${s.text.source}〕`,
          engine: 'pseudo',
          sourceLangUsed: s.language.value,
          targetLang: 'en-GB',
          at: new Date().toISOString(),
        };
      }
    }
    const out = await dtirToDocx(dtir, orig, { onMissingTranslation: 'error' });
    const xml = await docXml(out);

    expect(xml).toContain('EN〔Jaarverslag 2025〕');
    expect(xml).toContain('Resultaten'); // フィールドキャッシュは不可触
    expect(xml).toContain('1.250.000'); // 数値は不可触
    expect(xml).toContain('<w:sectPr'); // 構造保持
  });

  it('keeps original text when translation missing (keep mode)', async () => {
    const orig = readFileSync(fixture);
    const dtir = await docxToDtir(orig);
    const out = await dtirToDocx(dtir, orig, { onMissingTranslation: 'keep' });
    const xml = await docXml(out);
    expect(xml).toContain('Jaarverslag 2025'); // 未翻訳なら原文維持
  });

  it('collapses run-split paragraphs to a single text run', async () => {
    const orig = readFileSync(fixture);
    const dtir = await docxToDtir(orig);
    for (const s of dtir.segments) {
      if (s.translatable)
        s.translation = {
          text: 'X',
          engine: 'pseudo',
          sourceLangUsed: s.language.value,
          targetLang: 'en-GB',
          at: new Date().toISOString(),
        };
    }
    const out = await dtirToDocx(dtir, orig, { onMissingTranslation: 'keep' });
    const dtir2 = await docxToDtir(out);
    const fr = dtir2.segments.find((s) => (s.anchor.ref as { path: string }).path.includes('p[2]'));
    expect((fr?.anchor.ref as { runIds: string[] }).runIds.length).toBe(1);
  });
});
