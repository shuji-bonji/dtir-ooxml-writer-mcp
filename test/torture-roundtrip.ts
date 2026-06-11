/**
 * torture-roundtrip — 入れ子アンカー（表セル・脚注パート・ハイパーリンク/追跡変更内ラン）に
 * 対する writer のパッチ健全性を検証（reader 非依存・静的 torture DTIR を入力）。
 *
 * reader の再帰走査で richな anchor.ref.path（/w:body[1]/w:tbl[1]/w:tr[2]/w:tc[1]/w:p[1] や
 * /w:footnote[3]/w:p[1]）になった分、writer の汎用ナビゲータが正しく解決し、
 * **表セル・脚注・ハイパーリンク文に訳文を注入しつつ valid な Word docx を生成**できることを保証する。
 *
 * 実行: tsx test/torture-roundtrip.ts
 */
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import JSZip from 'jszip';
import { tortureDocxPath, tortureReaderDtirPath } from '@shuji-bonji/doc-translation-ir/fixtures';
import type { IRDocument } from '@shuji-bonji/doc-translation-ir';
import { dtirToDocx } from '../src/writer.js';

const GREEN = (s: string) => `\x1b[32m${s}\x1b[0m`;
const RED = (s: string) => `\x1b[31m${s}\x1b[0m`;
const mark = (s: string) => `EN〔${s}〕`;

async function partText(buf: Buffer, part: string): Promise<string> {
  const zip = await JSZip.loadAsync(buf);
  return (await zip.file(part)?.async('string')) ?? '';
}

async function main(): Promise<void> {
  const orig = readFileSync(tortureDocxPath);
  const dtir = JSON.parse(readFileSync(tortureReaderDtirPath, 'utf8')) as IRDocument;

  for (const seg of dtir.segments) {
    if (seg.translatable) {
      seg.translation = {
        text: mark(seg.text.source),
        engine: 'pseudo',
        sourceLangUsed: seg.language.value,
        targetLang: 'en-GB',
        at: new Date().toISOString(),
      };
    }
  }

  const out = await dtirToDocx(dtir, orig, { onMissingTranslation: 'error' });
  const docXml = await partText(out, 'word/document.xml');
  const footnotesXml = await partText(out, 'word/footnotes.xml');

  const failures: string[] = [];
  const ok = (c: boolean, m: string) => {
    if (!c) failures.push(m);
  };

  // 表セル（3つ・混在言語）への注入
  ok(docXml.includes(mark('Artikel 1')), '表セル(nl)の訳文が注入されていない');
  ok(docXml.includes(mark('Conditions générales')), '表セル(fr)の訳文が注入されていない');
  ok(docXml.includes(mark('Zusammenfassung der Bedingungen')), '結合セル(de)の訳文が注入されていない');
  // 表構造の保持
  ok(docXml.includes('<w:tbl'), 'w:tbl が失われた（表構造破壊）');
  ok(docXml.includes('<w:gridSpan'), 'gridSpan（結合セル）が失われた');

  // 脚注パートへの注入
  ok(footnotesXml.includes(mark('Vertrouwelijke voetnoot.')), '脚注の訳文が footnotes.xml に注入されていない');
  ok(/<w:footnote\b[^>]*w:type=/.test(footnotesXml), 'separator 脚注が失われた（脚注構造破壊）');

  // ハイパーリンク文 / 追跡変更文の collapse 注入（文全体が1本に集約）
  ok(docXml.includes(mark('See the signed contract for the full terms.')), 'ハイパーリンク文の訳文が注入されていない');
  ok(docXml.includes('<w:hyperlink'), 'w:hyperlink が失われた（リンク構造破壊）');
  ok(docXml.includes(mark('The deadline is strictly binding.')), '追跡変更文の訳文が注入されていない');
  ok(docXml.includes('<w:ins'), 'w:ins（追跡変更）が失われた');

  // valid zip
  ok((await JSZip.loadAsync(out)).file('word/document.xml') !== null, '出力が valid docx zip でない');

  // LibreOffice 変換（Word 互換）
  let pdfOk = false;
  try {
    const dir = mkdtempSync(join(tmpdir(), 'wa-torture-'));
    const docxPath = join(dir, 'out.docx');
    writeFileSync(docxPath, out);
    execFileSync('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', dir, docxPath], {
      stdio: 'ignore',
      timeout: 60000,
      env: { ...process.env, HOME: dir },
    });
    pdfOk = readFileSync(join(dir, 'out.pdf')).length > 0;
  } catch {
    pdfOk = false;
  }
  ok(pdfOk, 'LibreOffice が訳 torture docx を pdf 化できない');

  console.error('');
  if (failures.length === 0) {
    console.error(GREEN('TORTURE ROUNDTRIP PASS — 表/脚注/リンク/追跡変更へ訳注入・構造保持・Word互換'));
    process.exit(0);
  }
  console.error(RED(`FAIL — ${failures.length} 件`));
  for (const f of failures) console.error('  - ' + f);
  process.exit(1);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
