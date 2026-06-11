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

/**
 * 段落内の、テキスト(w:t)を持つランを**読み順で再帰収集**。
 * reader の collectRunEls と一致させる: w:hyperlink / w:ins / w:smartTag / w:sdt 内の
 * run も対象、w:del（削除済み）配下は除外。collapse の対象ランを過不足なく拾うため。
 */
function textRunsOf(p: El): El[] {
  const out: El[] = [];
  const walk = (el: El): void => {
    for (const c of childElements(el)) {
      if (c.tagName === 'w:del') continue;
      if (c.tagName === 'w:r') {
        if (descendantsByTag(c, 'w:t').length > 0) out.push(c);
        continue;
      }
      walk(c);
    }
  };
  walk(p);
  return out;
}

/**
 * part ルート（documentElement）からの構造パスをたどって要素を得る汎用ナビゲータ。
 * 各セグメントは `tag[idx]`（同名兄弟内 1 始まり）で、reader の collectParagraphs と一致する。
 * 例: /w:body[1]/w:tbl[1]/w:tr[2]/w:tc[1]/w:p[1]、/w:footnote[3]/w:p[1]、/w:p[1]
 */
function navigatePath(doc: Document, path: string): El | null {
  const segs = path.split('/').filter((s) => s.length > 0);
  let cur: El | null = doc.documentElement as unknown as El | null;
  for (const seg of segs) {
    if (!cur) return null;
    const m = /^([\w:]+)\[(\d+)\]$/.exec(seg);
    if (!m) return null;
    const tag = m[1];
    const idx = Number(m[2]);
    cur = childElements(cur).filter((e) => e.tagName === tag)[idx - 1] ?? null;
  }
  return cur;
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

/** ラン i の全 w:t のうち先頭へ text を入れ、残りを空にする（ラン内の書式は保持）。 */
function setRunText(run: El, doc: Document, text: string): void {
  let placed = false;
  for (const wt of descendantsByTag(run, 'w:t')) {
    setWtText(wt, doc, placed ? '' : text);
    placed = true;
  }
}

/**
 * 脱collapse: translation.runTexts を段落のテキストラン（読み順）へ 1:1 で分配する。
 * 各ランの rPr（太字・色・ハイパーリンク等）を保ったまま訳文を入れる。
 * runTexts の数が段落のテキストラン数と一致する場合のみ true（不一致なら呼び出し側が collapse）。
 */
function distributeRuns(p: El, doc: Document, runTexts: string[]): boolean {
  const runs = textRunsOf(p);
  if (runs.length !== runTexts.length) return false;
  runs.forEach((run, i) => setRunText(run, doc, runTexts[i]));
  return true;
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
      const p = navigatePath(doc, path);
      if (!p) throw new Error(`段落が見つからない: ${path} (${seg.id})`);
      if (p.tagName !== 'w:p') {
        throw new Error(`anchor.ref.path が w:p を指していない: ${path} (${seg.id})`);
      }
      const tr = seg.translation as { text: string; runTexts?: string[] };
      // runTexts があり段落のテキストラン数と一致すれば各ランへ分配（書式保持）、
      // 無い/不一致なら collapse（fail-safe）。
      const distributed = tr.runTexts ? distributeRuns(p, doc, tr.runTexts) : false;
      if (!distributed) collapseTranslation(p, doc, tr.text);
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
