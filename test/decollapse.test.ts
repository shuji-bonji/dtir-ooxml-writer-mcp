/**
 * vitest spec — 脱collapse（②）: translation.runTexts による各ランへの訳分配。
 * reader/translate 非依存（静的 torture DTIR に runTexts を手で詰めて writer を検証）。
 */
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { tortureDocxPath, tortureReaderDtirPath } from '@shuji-bonji/doc-translation-ir/fixtures';
import type { IRDocument } from '@shuji-bonji/doc-translation-ir';
import { dtirToDocx } from '../src/writer.js';

const orig = readFileSync(tortureDocxPath);
const loadDtir = (): IRDocument => JSON.parse(readFileSync(tortureReaderDtirPath, 'utf8')) as IRDocument;

async function docXml(buf: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(buf);
  return (await zip.file('word/document.xml')?.async('string')) ?? '';
}

/** 各ランを大文字化して runTexts を詰める（複数ラン段落のみ）。 */
function fillRunsUpper(dtir: IRDocument): IRDocument {
  for (const s of dtir.segments) {
    if (!s.translatable) continue;
    const base = { engine: 'pseudo', sourceLangUsed: s.language.value, targetLang: 'en-GB', at: new Date().toISOString() };
    if (s.text.runs && s.text.runs.length > 1) {
      const runTexts = s.text.runs.map((r) => s.text.source.slice(r.start, r.end).toUpperCase());
      s.translation = { text: runTexts.join(''), ...base, runTexts };
    } else {
      s.translation = { text: s.text.source.toUpperCase(), ...base };
    }
  }
  return dtir;
}

describe('dtirToDocx 脱collapse（runTexts 分配）', () => {
  it('段内の太字ランを保ったまま訳を分配する', async () => {
    const out = await dtirToDocx(fillRunsUpper(loadDtir()), orig, { onMissingTranslation: 'error' });
    const xml = await docXml(out);
    // 太字ランが <w:b/> を保ったまま自分の訳を持つ（collapse なら先頭ランに集約され太字が消える）
    expect(xml).toMatch(/<w:rPr><w:b\/><\/w:rPr><w:t[^>]*>MANDATORY<\/w:t>/);
    expect(xml).toContain('PAYMENT IS ');
    expect(xml).toContain(' NOW.');
  });

  it('ハイパーリンクの表示テキストを訳しつつリンク要素を保持', async () => {
    const out = await dtirToDocx(fillRunsUpper(loadDtir()), orig, { onMissingTranslation: 'error' });
    const xml = await docXml(out);
    expect(xml).toContain('<w:hyperlink');
    expect(xml).toContain('THE SIGNED CONTRACT');
  });

  it('runTexts の数が段落のラン数と合わなければ collapse にフォールバック（壊れない）', async () => {
    const dtir = loadDtir();
    const bold = dtir.segments.find((s) => s.text.source === 'Payment is mandatory now.')!;
    bold.translation = {
      text: 'X', engine: 'pseudo', sourceLangUsed: 'en-GB', targetLang: 'en-GB',
      at: new Date().toISOString(), runTexts: ['only-one'], // 段落は3ラン → 不一致
    };
    // 残りは通常訳で埋める
    for (const s of dtir.segments) {
      if (s.translatable && !s.translation) {
        s.translation = { text: s.text.source, engine: 'p', sourceLangUsed: s.language.value, targetLang: 'en-GB', at: new Date().toISOString() };
      }
    }
    const out = await dtirToDocx(dtir, orig, { onMissingTranslation: 'error' });
    const xml = await docXml(out);
    expect(xml).toContain('X'); // collapse で先頭ランに入る
    expect((await JSZip.loadAsync(out)).file('word/document.xml')).not.toBeNull();
  });
});
