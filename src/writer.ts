/**
 * writer — dtirToDocx: DTIR(翻訳済み) ＋ 元 .docx → 訳 .docx
 *
 * 設計（DTIR 契約に忠実）:
 *  - 元ファイルを基板に、anchor.ref の path で段落を特定し **id 単位でパッチ**。
 *  - translatable=true かつ translation!=null のセグメントのみ書き換える。
 *  - 非 translatable（フィールド/数値/空）と、IR に乗っていない要素
 *    （sectPr / DrawingML画像 / 書式）は**一切触らない**＝原理的に崩れない。
 *  - v0.1 は **collapse**: 訳文を先頭テキストランの w:t に入れ、残りの w:t は空にする
 *    （段内書式は捨てる）。段内書式保持は text.runs を使う v0.2 で対応。
 *
 * 「replace-by-id」は汎用の継ぎ目: 現状はテキスト part のパッチだが、将来
 *  バイナリ media part の差し替え（画像翻訳）も同じ枠組みで足せる。
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import JSZip from 'jszip';
import type { IRDocument, IRSegment } from '@shuji-bonji/doc-translation-ir';

const ELEMENT_NODE = 1;
const TEXT_NODE = 3;
type El = Element;

export interface DtirToDocxOptions {
  /** 翻訳が無いセグメントの扱い。'keep'=原文維持（既定）, 'error'=例外。 */
  onMissingTranslation?: 'keep' | 'error';
}

// --- DOM ヘルパ -------------------------------------------------------------
function childElements(el: El): El[] {
  const out: El[] = [];
  const nodes = el.childNodes;
  for (let i = 0; i < nodes.length; i++) {
    const n = nodes.item(i);
    if (n && n.nodeType === ELEMENT_NODE) out.push(n as unknown as El);
  }
  return out;
}

function descendantsByTag(el: El, tag: string): El[] {
  const list = el.getElementsByTagName(tag);
  const out: El[] = [];
  for (let i = 0; i < list.length; i++) out.push(list.item(i) as unknown as El);
  return out;
}

/** w:t 要素のテキストを置換（前後空白があれば xml:space=preserve を付与）。 */
function setWtText(wt: El, doc: Document, text: string): void {
  // 既存の子（テキストノード）を全削除
  while (wt.firstChild) wt.removeChild(wt.firstChild);
  if (text.length > 0) {
    wt.appendChild(doc.createTextNode(text));
    if (/^\s|\s$/.test(text)) wt.setAttribute('xml:space', 'preserve');
  }
}

/** 段落直下の、テキスト(w:t)を持つラン（読み順）。 */
function textRunsOf(p: El): El[] {
  return childElements(p)
    .filter((e) => e.tagName === 'w:r')
    .filter((r) => descendantsByTag(r, 'w:t').length > 0);
}

/** path 例: /w:body/w:p[12] → { container:'w:body', index:12 }。 */
function parsePath(path: string): { container: string; index: number } | null {
  const m = /^\/(w:\w+)\/w:p\[(\d+)\]$/.exec(path);
  if (!m) return null;
  return { container: m[1], index: Number(m[2]) };
}

/** collapse: 訳文を先頭 w:t に入れ、残り w:t を空にする。 */
function collapseTranslation(p: El, doc: Document, translation: string): void {
  const runs = textRunsOf(p);
  let placed = false;
  for (const run of runs) {
    for (const wt of descendantsByTag(run, 'w:t')) {
      if (!placed) {
        setWtText(wt, doc, translation);
        placed = true;
      } else {
        setWtText(wt, doc, '');
      }
    }
  }
}

// --- メイン -----------------------------------------------------------------
export async function dtirToDocx(
  dtir: IRDocument,
  originalDocx: Buffer | Uint8Array,
  options: DtirToDocxOptions = {},
): Promise<Buffer> {
  if (dtir.source.format !== 'docx') {
    throw new Error(`writer は docx 専用。dtir.source.format=${dtir.source.format}`);
  }
  const onMissing = options.onMissingTranslation ?? 'keep';
  const zip = await JSZip.loadAsync(originalDocx);

  // パッチ対象（translatable かつ translation あり）を part ごとに集約
  const byPart = new Map<string, IRSegment[]>();
  for (const seg of dtir.segments) {
    if (!seg.translatable) continue;
    if (!seg.translation) {
      if (onMissing === 'error') throw new Error(`未翻訳セグメント: ${seg.id}`);
      continue; // keep: 原文維持
    }
    const part = (seg.anchor.ref as { part?: string }).part;
    if (!part) throw new Error(`anchor.ref.part が無い: ${seg.id}`);
    const arr = byPart.get(part) ?? [];
    arr.push(seg);
    byPart.set(part, arr);
  }

  for (const [part, segs] of byPart) {
    const file = zip.file(part);
    if (!file) throw new Error(`part が元 docx に存在しない: ${part}`);
    const xml = await file.async('string');
    const doc = new DOMParser().parseFromString(xml, 'text/xml');

    for (const seg of segs) {
      const path = (seg.anchor.ref as { path?: string }).path ?? '';
      const parsed = parsePath(path);
      if (!parsed) throw new Error(`anchor.ref.path を解釈できない: ${path} (${seg.id})`);
      const container = doc.getElementsByTagName(parsed.container).item(0);
      if (!container) throw new Error(`container ${parsed.container} が無い: ${part}`);
      const paragraphs = childElements(container as unknown as El).filter(
        (e) => e.tagName === 'w:p',
      );
      const p = paragraphs[parsed.index - 1];
      if (!p) throw new Error(`段落が見つからない: ${path} (${seg.id})`);
      collapseTranslation(p, doc, (seg.translation as { text: string }).text);
    }

    let out = new XMLSerializer().serializeToString(doc);
    if (!out.startsWith('<?xml')) {
      out = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${out}`;
    }
    zip.file(part, out);
  }

  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  return buf;
}
