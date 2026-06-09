/**
 * roundtrip — dtir-ooxml-writer-mcp の受け入れテスト（reader と対）
 *
 *   fixture.docx --reader--> DTIR --擬似翻訳--> DTIR' --writer--> 訳.docx
 *
 * 検証:
 *  1. 訳文が translatable な w:t に注入されている
 *  2. 非 translatable（フィールドキャッシュ/数値）と構造（sectPr）は不可触
 *  3. ラン分断は collapse（先頭ランに集約、残りは空＝原文片が消える）
 *  4. 出力が valid zip / 全 XML well-formed（再 reader で確認）
 *  5. 再 reader 出力が validate-dtir を通る（構造保持）
 *  6. LibreOffice が訳 docx を pdf 化できる（Word 互換の妥当性）
 *
 * 実行: tsx test/roundtrip.ts
 */
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import Ajv2020 from 'ajv/dist/2020.js';
import addFormats from 'ajv-formats';
import JSZip from 'jszip';
import { docxToDtir } from '../../dtir-ooxml-reader-mcp/src/reader.js';
import { validateDtir } from '../../doc-translation-ir/tools/validate-dtir.js';
import { dtirToDocx } from '../src/writer.js';
import type { IRDocument } from '../src/dtir.js';

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, '..', '..');
const irRoot = resolve(repoRoot, 'doc-translation-ir');
const fixtureDocx = resolve(irRoot, 'fixtures/docx/mixed-nl-fr-de-tricky.docx');
const schemaPath = resolve(irRoot, 'schema/dtir-0.1.schema.json');

const GREEN = (s: string) => `\x1b[32m${s}\x1b[0m`;
const RED = (s: string) => `\x1b[31m${s}\x1b[0m`;
const mark = (s: string) => `EN〔${s}〕`;

async function partText(buf: Buffer, part: string): Promise<string> {
  const zip = await JSZip.loadAsync(buf);
  const f = zip.file(part);
  return f ? f.async('string') : '';
}

async function main(): Promise<void> {
  const orig = readFileSync(fixtureDocx);
  const failures: string[] = [];
  const ok = (cond: boolean, msg: string) => {
    if (!cond) failures.push(msg);
  };

  // reader → DTIR
  const dtir = await docxToDtir(orig, { fileName: 'mixed-nl-fr-de-tricky.docx', targetLang: 'en-GB' });

  // 擬似翻訳（translatable のみ）
  const expected = new Map<string, string>();
  for (const seg of dtir.segments) {
    if (!seg.translatable) continue;
    const text = mark(seg.text.source);
    expected.set(seg.id, text);
    seg.translation = {
      text,
      engine: 'pseudo',
      sourceLangUsed: seg.language.value,
      targetLang: 'en-GB',
      at: new Date().toISOString(),
    };
  }

  // writer → 訳 docx
  const out = await dtirToDocx(dtir, orig, { onMissingTranslation: 'error' });

  const docXml = await partText(out, 'word/document.xml');
  const headerXml = await partText(out, 'word/header1.xml');
  const footerXml = await partText(out, 'word/footer1.xml');

  // 1. 訳文注入
  ok(docXml.includes(mark('Jaarverslag 2025')), '見出しの訳文が注入されていない');
  ok(
    docXml.includes(mark('Les résultats du premier trimestre dépassent les prévisions.')),
    '仏語(ラン分断)の訳文が先頭ランに集約されていない',
  );
  ok(docXml.includes(mark('Die Produktion wurde im April vollständig automatisiert.')), '独語の訳文が注入されていない');
  ok(headerXml.includes(mark('Vertrouwelijk')), 'ヘッダの訳文が注入されていない');

  // 2. 不可触: フィールドキャッシュ・数値・構造
  ok(docXml.includes('Resultaten'), 'TOCフィールドのキャッシュが書き換えられた（不可触のはず）');
  ok(docXml.includes('1.250.000'), '数値が書き換えられた（不可触のはず）');
  ok(docXml.includes('<w:sectPr'), 'sectPr が失われた（構造破壊）');
  ok(/PAGE/.test(footerXml), 'フッタの PAGE フィールドが失われた');

  // 4+5. 再 reader で構造保持＋意味整合
  const dtir2 = await docxToDtir(out, { fileName: 'out.docx' });

  // 3. collapse（構造で検証）: 3ラン分断だった仏語段落が、訳注入後は
  //    テキストランが1本に集約されている（残りは空になり reader が除外）。
  const frBefore = dtir.segments.find((s) => s.text.source.includes('prévisions'));
  const frAfter = dtir2.segments.find((s) => s.text.source.includes('prévisions'));
  ok((frBefore?.anchor.ref as { runIds: string[] }).runIds.length === 3, '前提: 仏語は3ラン');
  ok(
    (frAfter?.anchor.ref as { runIds: string[] }).runIds.length === 1,
    `collapse 後の仏語が1ランに集約されていない（runs=${(frAfter?.anchor.ref as { runIds: string[] })?.runIds.length}）`,
  );
  const ajv = new Ajv2020({ allErrors: true, strict: false });
  addFormats(ajv);
  const validate = ajv.compile(JSON.parse(readFileSync(schemaPath, 'utf8')));
  ok(validate(dtir2) === true, '再 reader 出力が JSON Schema 不適合: ' + JSON.stringify(validate.errors?.slice(0, 2)));
  ok(validateDtir(dtir2).length === 0, '再 reader 出力が validate-dtir 不適合');
  const translatedSources = dtir2.segments.filter((s) => s.translatable).map((s) => s.text.source);
  ok(translatedSources.every((s) => s.startsWith('EN〔')), '再 reader で訳文が原文位置に来ていない');
  ok(dtir2.segments.length === dtir.segments.length, 'セグメント数が往復で変化（構造破壊）');

  // 6. LibreOffice 変換（Word 互換の妥当性）
  let pdfOk = false;
  try {
    const dir = mkdtempSync(join(tmpdir(), 'rt-'));
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
  ok(pdfOk, 'LibreOffice が訳 docx を pdf 化できない（妥当性 NG）');

  console.error('');
  if (failures.length === 0) {
    console.error(GREEN(`ROUNDTRIP ALL PASS — ${expected.size} 段落を翻訳注入、構造・フィールド・数値を保持、Word 互換確認`));
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
