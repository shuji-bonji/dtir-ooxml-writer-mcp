/**
 * writer acceptance — dtirToDocx を「静的 DTIR フィクスチャ」で検証（reader 非依存）
 *
 * 入力は doc-translation-ir 同梱の reader 出力 DTIR ＋ 元 docx。reader を実行しないので
 * このリポジトリは contract(doc-translation-ir) だけに依存する。
 * reader→translate→writer の本物 end-to-end は dtir-docx-pipeline リポジトリに在る。
 *
 * 検証: 訳文注入 / フィールド・数値・sectPr 不可触 / collapse（空ラン生成） /
 *       valid zip / LibreOffice で pdf 化（Word 互換）。
 *
 * 実行: tsx test/roundtrip.ts
 */
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import JSZip from 'jszip';
import { fixtureDocxPath, readerDtirPath } from '@shuji-bonji/doc-translation-ir/fixtures';
import type { IRDocument } from '@shuji-bonji/doc-translation-ir';
import { dtirToDocx } from '../src/writer.js';

const GREEN = (s: string) => `\x1b[32m${s}\x1b[0m`;
const RED = (s: string) => `\x1b[31m${s}\x1b[0m`;
const mark = (s: string) => `EN〔${s}〕`;
/** 空 w:t（self-closing か空タグ）の数＝collapse で潰れたラン数。 */
const emptyWtCount = (xml: string) =>
  (xml.match(/<w:t\b[^>]*\/>|<w:t\b[^>]*><\/w:t>/g) ?? []).length;

async function partText(buf: Buffer, part: string): Promise<string> {
  const zip = await JSZip.loadAsync(buf);
  return (await zip.file(part)?.async('string')) ?? '';
}

async function main(): Promise<void> {
  const orig = readFileSync(fixtureDocxPath);
  const dtir = JSON.parse(readFileSync(readerDtirPath, 'utf8')) as IRDocument;

  // 擬似翻訳（translatable のみ）
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
  const headerXml = await partText(out, 'word/header1.xml');
  const footerXml = await partText(out, 'word/footer1.xml');

  const failures: string[] = [];
  const ok = (c: boolean, m: string) => {
    if (!c) failures.push(m);
  };

  // 訳文注入
  ok(docXml.includes(mark('Jaarverslag 2025')), '見出しの訳文が注入されていない');
  ok(
    docXml.includes(mark('Les résultats du premier trimestre dépassent les prévisions.')),
    '仏語(3ラン)の訳文が先頭ランに集約されていない',
  );
  ok(docXml.includes(mark('Die Produktion wurde im April vollständig automatisiert.')), '独語の訳文が注入されていない');
  ok(headerXml.includes(mark('Vertrouwelijk')), 'ヘッダの訳文が注入されていない');

  // 不可触
  ok(docXml.includes('Resultaten'), 'TOCフィールドのキャッシュが書き換えられた');
  ok(docXml.includes('1.250.000'), '数値が書き換えられた');
  ok(docXml.includes('<w:sectPr'), 'sectPr が失われた（構造破壊）');
  ok(/PAGE/.test(footerXml), 'フッタの PAGE フィールドが失われた');

  // collapse: 仏語の3ラン→先頭1本に集約し、残り2ランの w:t が空になる
  ok(emptyWtCount(docXml) >= 2, `collapse で空ランが生成されていない（空 w:t=${emptyWtCount(docXml)}）`);

  // valid zip
  ok((await JSZip.loadAsync(out)).file('word/document.xml') !== null, '出力が valid docx zip でない');

  // LibreOffice 変換（Word 互換）
  let pdfOk = false;
  try {
    const dir = mkdtempSync(join(tmpdir(), 'wa-'));
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
  ok(pdfOk, 'LibreOffice が訳 docx を pdf 化できない');

  console.error('');
  if (failures.length === 0) {
    console.error(GREEN('WRITER ACCEPTANCE PASS — 静的DTIRから訳文注入・構造保持・collapse・Word互換を確認'));
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
