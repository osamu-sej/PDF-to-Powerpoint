// =========================================================
// refine.js — 「うちた式パワポ変換 V1.5」の要件を変換後の PPTX に適用する後処理
//
// LibreOffice の PDF インポートは「線＝独立した prstGeom=line の図形」「矢尻＝別の塗り三角形」
// 「四角形＝custGeom の自由曲線」として出力する。そのままでは図形を動かしても線が追従せず、
// 1本の矢印が 2 個のオブジェクトに分かれて編集しづらい。ここでは次の処理を行う。
//
//   1. 負のサイズ（ext < 0）を flipH/flipV に正規化
//   2. 軸平行の四角形 custGeom を native な rect に戻す（コネクタ接着の前提）
//   3. 短線の連なりで作られた破線を、1 本の破線（線スタイル）に統合
//   4. 別図形の三角矢尻を、線の headEnd/tailEnd に統合（矢印の向きを終点側に揃える）
//   5. 直線を PowerPoint のコネクタ（p:cxnSp）にし、ノード（rect）の接続点へ接着
//   6. custDash を dash / sysDash / sysDot の線スタイルへ寄せる（近い場合のみ）
//   7. eastAsia と latin に同じフェイスを指定し、非標準フォントは標準フォントへ置換
//   8. 透過 TextBox の右マージンを広げる（wrap=none は既存処理で設定済み）
//
// 併せて、出力 PPTX を読み戻して 6 項目を数える機械検品（auditPptx）と、
// 納品メモ（buildDeliveryNote）を生成する。
// 解析に失敗した場合は元の PPTX を一切変更しない（呼び出し側で握りつぶす）。
// =========================================================
import { createRequire } from 'module';
import fs from 'fs';
const require = createRequire(import.meta.url);
const JSZip = require('jszip');

export const EMU_PER_PT = 12700;
const PT = EMU_PER_PT;

// ---------------------------------------------------------
// 設定値（V1.5 の許容値に合わせる）
// ---------------------------------------------------------
export const LANDING_TOL_PT = 2;      // 終点が辺上にあるとみなす許容（±2pt）
const GLUE_TOL_PT = 4;                // 接続点（辺の中点）に吸着させる許容
const TEXTBOX_RIGHT_MARGIN_PT = 30;   // TextBox の右マージン（30〜100pt の下限）
const ARROW_MAX_PT = 24;              // これを超える三角形は矢尻とみなさない
const NODE_MIN_PT = 8;                // これ未満の図形はノードとみなさない
const DASH_PRESET_MAX_DIST = 0.06;    // custDash を線スタイルへ寄せる許容（対数距離の二乗和。見た目が変わるほど遠い場合は custDash のまま保持）

// 閲覧側の PowerPoint / Windows / Office に標準搭載されているフォント（正規化キー）
const STANDARD_FONT_KEYS = new Set([
    'yugothic', 'yumincho', 'meiryo', 'meiryoui', 'msgothic', 'mspgothic', 'msmincho', 'mspmincho',
    'ｍｓゴシック', 'ｍｓｐゴシック', 'ｍｓ明朝', 'ｍｓｐ明朝',
    'arial', 'arialnarrow', 'arialblack', 'timesnewroman', 'couriernew', 'calibri', 'calibrilight',
    'cambria', 'segoeui', 'segoeuisemibold', 'segoeuilight', 'verdana', 'tahoma', 'georgia',
    'consolas', 'centurygothic', 'trebuchetms', 'candara', 'garamond', 'symbol', 'wingdings',
    'bizudgothic', 'bizudpgothic', 'bizudmincho', 'bizudpmincho',
    'bizudminchomedium', 'bizudpminchomedium',
    'udデジタル教科書体nr', 'udデジタル教科書体npr', 'udデジタル教科書体nk',
].map(k => k.toLowerCase()));

const fontKey = (n) => String(n || '').toLowerCase().replace(/[\s\-_,]/g, '');
export function isStandardFont(name) {
    if (!name) return true;
    if (name.startsWith('+')) return true; // テーマ参照（+mn-lt 等）
    const k = fontKey(name);
    if (STANDARD_FONT_KEYS.has(k)) return true;
    return /^udデジタル教科書体/.test(k) || /^bizud/.test(k);
}

// ---------------------------------------------------------
// 小物
// ---------------------------------------------------------
const decodeEntities = (s) => s.replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&#(\d+);/g, (m, d) => String.fromCodePoint(parseInt(d, 10))).replace(/&amp;/g, '&');
const escapeAttr = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
const hasCjk = (s) => /[\u3000-\u30ff\u3400-\u9fff\uf900-\ufaff\uff00-\uffef]/.test(s);
const dist = (a, b) => Math.hypot(a.x - b.x, a.y - b.y);
const round = (n) => Math.round(n);

// 文字幅の推定（V1.5 検品⑥の基準: 全角＝サイズ×1.0／半角＝×0.56）
export function estimateTextWidthPt(text, sizePt) {
    let w = 0;
    for (const ch of text) {
        const c = ch.codePointAt(0);
        if (c === 0x20) w += 0.56 * sizePt * 0.6;
        else if (c < 0x0100 || (c >= 0xff61 && c <= 0xff9f)) w += 0.56 * sizePt;
        else w += sizePt;
    }
    return w;
}

// ---------------------------------------------------------
// spTree の直下要素の分解／再結合
// ---------------------------------------------------------
const TOP_TAGS = new Set(['sp', 'pic', 'cxnSp', 'grpSp', 'graphicFrame', 'contentPart']);

export function splitSpTree(xml) {
    const open = xml.indexOf('<p:spTree>');
    const close = xml.lastIndexOf('</p:spTree>');
    if (open < 0 || close < 0) return null;
    const grpEnd = xml.indexOf('</p:grpSpPr>', open);
    if (grpEnd < 0) return null;
    const head = xml.slice(0, grpEnd + '</p:grpSpPr>'.length);
    const body = xml.slice(grpEnd + '</p:grpSpPr>'.length, close);
    const tail = xml.slice(close);
    const items = [];
    let pos = 0;
    while (pos < body.length) {
        const m = /<p:([A-Za-z]+)\b/.exec(body.slice(pos));
        if (!m) { items.push({ tag: 'raw', xml: body.slice(pos) }); break; }
        const start = pos + m.index;
        const name = m[1];
        if (start > pos) items.push({ tag: 'raw', xml: body.slice(pos, start) });
        let end;
        if (!TOP_TAGS.has(name)) {
            // 未知の要素（extLst 等）は対応する閉じタグまでを不透明な塊として扱う
            const c = body.indexOf(`</p:${name}>`, start);
            end = c < 0 ? body.length : c + `</p:${name}>`.length;
            items.push({ tag: 'raw', xml: body.slice(start, end) });
            pos = end;
            continue;
        }
        // 同名タグの入れ子（grpSp）を数えて閉じタグを探す
        const openRe = new RegExp(`<p:${name}[\\s>]`, 'g');
        const closeTag = `</p:${name}>`;
        let depth = 0;
        let scan = start;
        end = -1;
        while (scan < body.length) {
            openRe.lastIndex = scan;
            const o = openRe.exec(body);
            const c = body.indexOf(closeTag, scan);
            if (c < 0) break;
            if (o && o.index < c) { depth++; scan = o.index + 1; }
            else { depth--; scan = c + closeTag.length; if (depth === 0) { end = scan; break; } }
        }
        if (end < 0) return null;
        items.push({ tag: name, xml: body.slice(start, end) });
        pos = end;
    }
    return { head, items, tail };
}

export const joinSpTree = (t) => t.head + t.items.map(i => i.xml).join('') + t.tail;

// ---------------------------------------------------------
// 図形情報の抽出
// ---------------------------------------------------------
const XFRM_RE = /<a:xfrm((?:\s+[a-zA-Z]+="[^"]*")*)\s*>\s*<a:off x="(-?\d+)" y="(-?\d+)"\s*\/>\s*<a:ext cx="(-?\d+)" cy="(-?\d+)"\s*\/>/;

function parseXfrm(xml) {
    const m = XFRM_RE.exec(xml);
    if (!m) return null;
    const attrs = m[1] || '';
    return {
        match: m[0], index: m.index,
        flipH: /flipH="1"/.test(attrs), flipV: /flipV="1"/.test(attrs),
        rot: parseInt((attrs.match(/\brot="(-?\d+)"/) || [0, 0])[1], 10),
        x: parseInt(m[2], 10), y: parseInt(m[3], 10),
        cx: parseInt(m[4], 10), cy: parseInt(m[5], 10),
    };
}

function buildXfrmTag(x, y, cx, cy, flipH, flipV, rot = 0) {
    const a = `${rot ? ` rot="${rot}"` : ''}${flipH ? ' flipH="1"' : ''}${flipV ? ' flipV="1"' : ''}`;
    return `<a:xfrm${a}><a:off x="${round(x)}" y="${round(y)}"/><a:ext cx="${round(cx)}" cy="${round(cy)}"/>`;
}

const GEOM_RE = /<a:custGeom>[\s\S]*?<\/a:custGeom>|<a:prstGeom prst="[^"]*"\s*(?:\/>|>[\s\S]*?<\/a:prstGeom>)/;

// custGeom のパス（moveTo / lnTo / close のみ対応。曲線を含む場合は null）
function parseCustPath(xml) {
    const g = /<a:custGeom>([\s\S]*?)<\/a:custGeom>/.exec(xml);
    if (!g) return null;
    const body = g[1];
    const paths = body.match(/<a:path\b[\s\S]*?<\/a:path>/g) || [];
    if (paths.length !== 1) return null;
    if (/<a:(cubicBezTo|quadBezTo|arcTo)\b/.test(paths[0])) return null;
    const head = /<a:path\b([^>]*)>/.exec(paths[0])[1];
    const w = parseInt((head.match(/\bw="(\d+)"/) || [0, 0])[1], 10);
    const h = parseInt((head.match(/\bh="(\d+)"/) || [0, 0])[1], 10);
    const pts = [];
    let closed = false;
    const re = /<a:(moveTo|lnTo|close)\b[^>]*?(?:\/>|>([\s\S]*?)<\/a:\1>)/g;
    let m;
    while ((m = re.exec(paths[0]))) {
        if (m[1] === 'close') { closed = true; continue; }
        const p = /<a:pt x="(-?\d+)" y="(-?\d+)"/.exec(m[2] || '');
        if (!p) return null;
        pts.push({ cmd: m[1], x: parseInt(p[1], 10), y: parseInt(p[2], 10) });
    }
    if (!pts.length || pts[0].cmd !== 'moveTo' || pts.slice(1).some(p => p.cmd !== 'lnTo')) return null;
    return { w, h, pts, closed };
}

function parseLn(spPr) {
    const m = /<a:ln\b([^>]*?)(?:\/>|>([\s\S]*?)<\/a:ln>)/.exec(spPr);
    if (!m) return { present: false, noFill: false, w: 0, color: null, dash: null, custDash: null, head: null, tail: null, match: null };
    const attrs = m[1] || '';
    const body = m[2] || '';
    const w = parseInt((attrs.match(/\bw="(\d+)"/) || [0, 0])[1], 10);
    const noFill = /<a:noFill\s*\/>/.test(body);
    const color = ((/<a:solidFill>\s*<a:srgbClr val="([0-9a-fA-F]{6})"/.exec(body) || [])[1] || '').toLowerCase() || null;
    const dash = (/<a:prstDash val="([^"]+)"/.exec(body) || [])[1] || null;
    const cd = /<a:custDash>\s*<a:ds d="(\d+)" sp="(\d+)"/.exec(body);
    const ends = (tag) => {
        const e = new RegExp(`<a:${tag}\\b([^>]*)/?>`).exec(body);
        if (!e) return null;
        return { type: (e[1].match(/type="([^"]+)"/) || [])[1] || 'none' };
    };
    return {
        present: true, noFill, w, color, dash,
        custDash: cd ? { d: parseInt(cd[1], 10), sp: parseInt(cd[2], 10) } : null,
        head: ends('headEnd'), tail: ends('tailEnd'),
        match: m[0],
    };
}

function parseFill(spPr) {
    const noLn = spPr.replace(/<a:ln\b[^>]*?(?:\/>|>[\s\S]*?<\/a:ln>)/, '');
    if (/<a:noFill\s*\/>/.test(noLn)) return { kind: 'none', color: null };
    const m = /<a:solidFill>\s*<a:srgbClr val="([0-9a-fA-F]{6})"/.exec(noLn);
    if (m) return { kind: 'solid', color: m[1].toLowerCase() };
    if (/<a:(gradFill|blipFill|pattFill)\b/.test(noLn)) return { kind: 'other', color: null };
    return { kind: 'none', color: null };
}

// 図形（p:sp / p:cxnSp）の解析
export function parseShape(item) {
    if (item.tag !== 'sp' && item.tag !== 'cxnSp') return null;
    const xml = item.xml;
    const id = parseInt((/<p:cNvPr id="(\d+)"/.exec(xml) || [0, 0])[1], 10);
    const xf = parseXfrm(xml);
    if (!xf) return null;
    const spPr = (/<p:spPr\b[\s\S]*?<\/p:spPr>/.exec(xml) || [''])[0];
    const prst = (/<a:prstGeom prst="([^"]+)"/.exec(spPr) || [])[1] || null;
    const cust = prst ? null : parseCustPath(spPr);
    const hasCust = /<a:custGeom>/.test(spPr);
    const ln = parseLn(spPr);
    const fill = parseFill(spPr);
    const textRuns = [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map(m => decodeEntities(m[1]));
    const hasText = textRuns.some(t => t.trim() !== '');
    const isTxBox = /<p:cNvSpPr[^>]*\btxBox="1"/.test(xml);
    const info = {
        item, id, xf, prst, cust, hasCust, ln, fill, hasText, isTxBox, spPr,
        lnVisible: ln.present && !ln.noFill && ln.w > 0 && !!ln.color,
        kind: 'other',
    };
    // 種別判定
    if (item.tag === 'cxnSp') info.kind = 'connector';
    else if (prst === 'line' || prst === 'straightConnector1') info.kind = 'line';
    else if (prst === 'rect' || prst === 'roundRect' || /^flowChart(Process|Alternate|Document|Terminator|Decision|PredefinedProcess)/.test(prst || '')) info.kind = 'rectlike';
    else if (prst === 'ellipse') info.kind = 'ellipse';
    else if (cust) info.kind = classifyCust(info);
    return info;
}

// 絶対座標（EMU）の点列へ展開（flip 反映）
function absPoints(info) {
    const { xf } = info;
    const tf = (px, py, pw, ph) => {
        let X = pw ? px / pw * xf.cx : 0;
        let Y = ph ? py / ph * xf.cy : 0;
        if (xf.flipH) X = xf.cx - X;
        if (xf.flipV) Y = xf.cy - Y;
        return { x: xf.x + X, y: xf.y + Y };
    };
    if (info.prst === 'line' || info.prst === 'straightConnector1') {
        const a = tf(0, 0, 1, 1), b = tf(1, 1, 1, 1);
        return [a, b];
    }
    if (info.cust) return info.cust.pts.map(p => tf(p.x, p.y, info.cust.w, info.cust.h));
    return null;
}

function uniqueVerts(pts) {
    const out = [];
    for (const p of pts) if (!out.some(q => Math.abs(q.x - p.x) < 2 && Math.abs(q.y - p.y) < 2)) out.push(p);
    return out;
}

function classifyCust(info) {
    const { cust, fill } = info;
    const pts = absPoints(info);
    if (!pts) return 'other';
    const uv = uniqueVerts(pts);
    if (!cust.closed && fill.kind === 'none' && uv.length >= 2) return uv.length === 2 ? 'line' : 'polyline';
    if (cust.closed && uv.length === 3) return 'triangle';
    if (cust.closed && uv.length === 4) {
        const xs = [...new Set(uv.map(p => Math.round(p.x / 4)))];
        const ys = [...new Set(uv.map(p => Math.round(p.y / 4)))];
        if (xs.length === 2 && ys.length === 2) return 'rectlike';
    }
    return 'other';
}

// 線（2 点）か折れ線の点列を返す。それ以外は null
function linePoints(info) {
    if (info.kind === 'line' || info.kind === 'connector') {
        const p = absPoints(info);
        return p && p.length >= 2 ? [p[0], p[p.length - 1]] : null;
    }
    if (info.kind === 'polyline') return uniqueVerts(absPoints(info));
    return null;
}

// ---------------------------------------------------------
// 1. 負サイズの正規化
// ---------------------------------------------------------
export function fixNegativeSizes(xml) {
    let fixed = 0;
    const out = xml.replace(/<a:xfrm((?:\s+[a-zA-Z]+="[^"]*")*)\s*>(\s*)<a:off x="(-?\d+)" y="(-?\d+)"\s*\/>(\s*)<a:ext cx="(-?\d+)" cy="(-?\d+)"\s*\/>/g,
        (m, attrs, s1, x, y, s2, cx, cy) => {
            let X = parseInt(x, 10), Y = parseInt(y, 10), CX = parseInt(cx, 10), CY = parseInt(cy, 10);
            if (CX >= 0 && CY >= 0) return m;
            fixed++;
            let flipH = /flipH="1"/.test(attrs), flipV = /flipV="1"/.test(attrs);
            if (CX < 0) { X += CX; CX = -CX; flipH = !flipH; }
            if (CY < 0) { Y += CY; CY = -CY; flipV = !flipV; }
            let a = attrs.replace(/\s+flip[HV]="[^"]*"/g, '');
            if (flipH) a += ' flipH="1"';
            if (flipV) a += ' flipV="1"';
            return `<a:xfrm${a}>${s1}<a:off x="${X}" y="${Y}"/>${s2}<a:ext cx="${CX}" cy="${CY}"/>`;
        });
    return { xml: out, fixed };
}

// ---------------------------------------------------------
// 7. eastAsia / latin の揃え、非標準フォントの置換
// ---------------------------------------------------------
// run 単位でテキストを見て CJK / 欧文を判定し、latin と ea に同じフェイスを指定する
export function normalizeFonts(xml, { forceFont = null } = {}) {
    let paired = 0, replaced = 0;
    const FALLBACK_CJK = 'Yu Gothic';
    const FALLBACK_LATIN = 'Arial';

    const fixProps = (tag, attrs, inner, text) => {
        const lat = /<a:latin typeface="([^"]*)"([^>]*?)\/>/.exec(inner);
        const ea = /<a:ea typeface="([^"]*)"([^>]*?)\/>/.exec(inner);
        if (!lat && !ea) return null;
        if ((lat && lat[1].startsWith('+')) || (ea && ea[1].startsWith('+'))) return null; // テーマ参照は触らない
        const latinFace = lat ? decodeEntities(lat[1]) : null;
        const eaFace = ea ? decodeEntities(ea[1]) : null;
        const cjk = text != null ? hasCjk(text) : null;
        let face;
        if (forceFont) face = forceFont;
        else if (latinFace && eaFace) face = (cjk === false) ? latinFace : eaFace;
        else face = eaFace || latinFace;
        if (!forceFont && !isStandardFont(face)) {
            replaced++;
            face = cjk === false ? FALLBACK_LATIN : (cjk ? FALLBACK_CJK : (eaFace && !latinFace ? FALLBACK_CJK : FALLBACK_LATIN));
        }
        const esc = escapeAttr(face);
        let body = inner;
        // 既存の latin / ea を除き、スキーマ順（latin → ea → cs）で直後に挿入し直す
        body = body.replace(/<a:latin\b[^>]*?\/>/g, '').replace(/<a:ea\b[^>]*?\/>/g, '');
        const pair = `<a:latin typeface="${esc}"/><a:ea typeface="${esc}"/>`;
        const csM = /<a:cs\b[^>]*?\/>/.exec(body);
        const anchor = /<a:(sym|hlinkClick|hlinkMouseOver|rtl|extLst)\b/.exec(body);
        if (csM) body = body.slice(0, csM.index) + pair + body.slice(csM.index);
        else if (anchor) body = body.slice(0, anchor.index) + pair + body.slice(anchor.index);
        else body += pair;
        if (!(lat && ea && latinFace === eaFace && latinFace === face)) paired++;
        return `<a:${tag}${attrs}>${body}</a:${tag}>`;
    };

    // run（テキスト付き）。自己終了の rPr は対象外（次の閉じタグまで食い込まないよう明示的に分岐）
    let out = xml.replace(/<a:r>(\s*)<a:rPr\b([^>]*?)(\/>|>([\s\S]*?)<\/a:rPr>)(\s*)<a:t>([\s\S]*?)<\/a:t>/g,
        (m, s1, attrs, rest, inner, s2, t) => {
            if (rest === '/>') return m;
            const r = fixProps('rPr', attrs, inner, decodeEntities(t));
            return r ? `<a:r>${s1}${r}${s2}<a:t>${t}</a:t>` : m;
        });
    // endParaRPr / defRPr（テキスト不明）
    out = out.replace(/<a:(endParaRPr|defRPr)\b([^>]*?)(\/>|>([\s\S]*?)<\/a:\1>)/g, (m, tag, attrs, rest, inner) => {
        if (rest === '/>') return m;
        const r = fixProps(tag, attrs, inner, null);
        return r || m;
    });
    return { xml: out, paired, replaced };
}

// ---------------------------------------------------------
// 6. custDash → 線スタイル
// ---------------------------------------------------------
const DASH_PRESETS = { sysDot: [100, 100], sysDash: [300, 100], dash: [400, 300] };

export function nearestDashPreset(dPct, spPct) {
    let best = null;
    for (const [name, [d, sp]] of Object.entries(DASH_PRESETS)) {
        const dd = Math.log(dPct / d), ds = Math.log(spPct / sp);
        const dist2 = dd * dd + ds * ds;
        if (!best || dist2 < best.dist2) best = { name, dist2 };
    }
    return best;
}

export function normalizeDashStyles(xml) {
    let toPreset = 0, kept = 0;
    const out = xml.replace(/<a:custDash>\s*<a:ds d="(\d+)" sp="(\d+)"\s*\/>\s*<\/a:custDash>/g, (m, d, sp) => {
        const dPct = parseInt(d, 10) / 1000, spPct = parseInt(sp, 10) / 1000;
        if (dPct <= 0 || spPct <= 0) { kept++; return m; }
        const b = nearestDashPreset(dPct, spPct);
        if (b.dist2 <= DASH_PRESET_MAX_DIST) { toPreset++; return `<a:prstDash val="${b.name}"/>`; }
        kept++;
        return m;
    });
    return { xml: out, toPreset, kept };
}

// ---------------------------------------------------------
// 2〜5. 図形レベルの再構築
// ---------------------------------------------------------
function nodeEdgeDistance(node, p) {
    // node: {x,y,cx,cy,kind}; p: 点 → 輪郭までの距離（EMU）
    const { x, y, cx, cy } = node;
    if (node.kind === 'ellipse') {
        const a = cx / 2, b = cy / 2;
        const nx = (p.x - (x + a)) / a, ny = (p.y - (y + b)) / b;
        return Math.abs(Math.hypot(nx, ny) - 1) * Math.min(a, b);
    }
    const dx = Math.max(x - p.x, 0, p.x - (x + cx));
    const dy = Math.max(y - p.y, 0, p.y - (y + cy));
    if (dx > 0 || dy > 0) return Math.hypot(dx, dy); // 外側
    return Math.min(p.x - x, x + cx - p.x, p.y - y, y + cy - p.y); // 内側
}

function nodeSites(node) {
    const { x, y, cx, cy } = node;
    // rect / roundRect の接続点: 0=上, 1=左, 2=下, 3=右
    return [
        { idx: 0, x: x + cx / 2, y },
        { idx: 1, x, y: y + cy / 2 },
        { idx: 2, x: x + cx / 2, y: y + cy },
        { idx: 3, x: x + cx, y: y + cy / 2 },
    ];
}

function collectNodes(infos, slideW, slideH) {
    const nodes = [];
    for (const s of infos) {
        if (!s) continue;
        if (!['rectlike', 'ellipse'].includes(s.kind)) continue;
        if (s.isTxBox) continue;
        if (s.xf.rot) continue;
        const visible = s.fill.kind !== 'none' || s.lnVisible;
        if (!visible) continue;
        if (s.xf.cx < NODE_MIN_PT * PT || s.xf.cy < NODE_MIN_PT * PT) continue;
        // スライド全面の背景・ほぼ全面の枠は接着先にしない
        if (s.xf.cx * s.xf.cy >= 0.9 * slideW * slideH) continue;
        nodes.push({ id: s.id, x: s.xf.x, y: s.xf.y, cx: s.xf.cx, cy: s.xf.cy, kind: s.kind === 'ellipse' ? 'ellipse' : 'rect', prst: s.prst });
    }
    return nodes;
}

// 三角形（矢尻）の幾何
function triangleGeom(info) {
    const pts = uniqueVerts(absPoints(info));
    if (pts.length !== 3) return null;
    // 先端 = 隣接 2 辺の長さがもっとも近い頂点（二等辺三角形の頂角）
    let tipIdx = 0, bestRatio = Infinity;
    for (let i = 0; i < 3; i++) {
        const a = dist(pts[i], pts[(i + 1) % 3]), b = dist(pts[i], pts[(i + 2) % 3]);
        const r = Math.max(a, b) / Math.max(1, Math.min(a, b));
        if (r < bestRatio) { bestRatio = r; tipIdx = i; }
    }
    if (bestRatio > 1.35) return null; // 二等辺とみなせない
    const tip = pts[tipIdx], b1 = pts[(tipIdx + 1) % 3], b2 = pts[(tipIdx + 2) % 3];
    const baseMid = { x: (b1.x + b2.x) / 2, y: (b1.y + b2.y) / 2 };
    const axisLen = dist(tip, baseMid);
    const baseW = dist(b1, b2);
    if (axisLen < 1 || axisLen > ARROW_MAX_PT * PT || baseW > ARROW_MAX_PT * PT) return null;
    return { tip, baseMid, axisLen, baseW, dir: { x: (tip.x - baseMid.x) / axisLen, y: (tip.y - baseMid.y) / axisLen } };
}

function arrowSize(ratio) { // 線幅に対する比 → sm / med / lg
    // PowerPoint の矢印サイズは線幅の 2 / 3 / 5 倍。幾何平均を境に最も近いものを選ぶ
    if (ratio < 2.45) return 'sm';
    if (ratio < 3.87) return 'med';
    return 'lg';
}

// ln 要素へ headEnd / tailEnd を設定する
function setLineEnds(xml, { head = null, tail = null }) {
    const mk = (tag, e) => e ? `<a:${tag} type="${e.type}" w="${e.w}" len="${e.len}"/>` : '';
    return xml.replace(/<a:ln\b([^>]*?)(\/>|>([\s\S]*?)<\/a:ln>)/, (m, attrs, rest, inner) => {
        let body = rest === '/>' ? '' : inner;
        if (head) body = body.replace(/<a:headEnd\b[^>]*?\/>/g, '');
        if (tail) body = body.replace(/<a:tailEnd\b[^>]*?\/>/g, '');
        // スキーマ順: fill → dash → join → headEnd → tailEnd
        const ext = body.match(/<a:extLst>[\s\S]*<\/a:extLst>/);
        let core = ext ? body.replace(ext[0], '') : body;
        core += mk('headEnd', head) + mk('tailEnd', tail);
        return `<a:ln${attrs}>${core}${ext ? ext[0] : ''}</a:ln>`;
    });
}

function lineGeomXml(p0, p1) {
    const x = Math.min(p0.x, p1.x), y = Math.min(p0.y, p1.y);
    return {
        x, y, cx: Math.abs(p1.x - p0.x), cy: Math.abs(p1.y - p0.y),
        flipH: p0.x > p1.x, flipV: p0.y > p1.y,
    };
}

// 2 点の直線 p:sp → p:cxnSp（直線コネクタ）
function toConnector(info, p0, p1, st, en) {
    const g = lineGeomXml(p0, p1);
    const name = `コネクタ ${info.id}`;
    const cxn = (st ? `<a:stCxn id="${st.id}" idx="${st.idx}"/>` : '') + (en ? `<a:endCxn id="${en.id}" idx="${en.idx}"/>` : '');
    const ln = info.ln.match || '<a:ln/>';
    return `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${info.id}" name="${name}"/><p:cNvCxnSpPr>${cxn}</p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>` +
        `<p:spPr>${buildXfrmTag(g.x, g.y, g.cx, g.cy, g.flipH, g.flipV)}</a:xfrm><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>${ln}</p:spPr></p:cxnSp>`;
}

// 折れ線 custGeom を点列から作り直す（先端を矢尻の先端へ移した場合など）
function rebuildPolylineXml(xml, pts) {
    const minX = Math.min(...pts.map(p => p.x)), minY = Math.min(...pts.map(p => p.y));
    const maxX = Math.max(...pts.map(p => p.x)), maxY = Math.max(...pts.map(p => p.y));
    const cx = Math.max(1, round(maxX - minX)), cy = Math.max(1, round(maxY - minY));
    const rel = pts.map(p => ({ x: round(p.x - minX), y: round(p.y - minY) }));
    const path = `<a:path w="${cx}" h="${cy}">` +
        rel.map((p, i) => `<a:${i ? 'lnTo' : 'moveTo'}><a:pt x="${p.x}" y="${p.y}"/></a:${i ? 'lnTo' : 'moveTo'}>`).join('') + '</a:path>';
    const cust = `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>${path}</a:pathLst></a:custGeom>`;
    let out = xml.replace(XFRM_RE, buildXfrmTag(minX, minY, cx, cy, false, false) );
    // buildXfrmTag は開始タグ＋off/ext までを返すので、元の閉じタグ位置を維持する
    out = out.replace(GEOM_RE, cust);
    return out;
}

// 短線の連なり（自作の破線）を 1 本の破線へ統合する
function mergeDashedSegments(infos, stats) {
    const segs = [];
    infos.forEach((s, idx) => {
        if (!s || s.kind !== 'line' || s.item.tag !== 'sp' || s.hasText || !s.lnVisible) return;
        if (s.ln.dash || s.ln.custDash || s.ln.head || s.ln.tail) return;
        const p = linePoints(s);
        if (!p) return;
        const len = dist(p[0], p[1]);
        if (len < 0.5 * PT) return;
        segs.push({ idx, s, p0: p[0], p1: p[1], len, ang: Math.atan2(p[1].y - p[0].y, p[1].x - p[0].x) });
    });
    const groups = new Map();
    for (const g of segs) {
        let a = g.ang; if (a < 0) a += Math.PI; if (a >= Math.PI - 1e-6) a = 0;
        const nx = -Math.sin(a), ny = Math.cos(a);
        const off = (g.p0.x * nx + g.p0.y * ny) / PT;
        const key = `${g.s.ln.color}|${g.s.ln.w}|${Math.round(a * 180 / Math.PI)}|${Math.round(off / 0.75)}`;
        if (!groups.has(key)) groups.set(key, []);
        groups.get(key).push({ ...g, a });
    }
    const removed = new Set();
    for (const list of groups.values()) {
        if (list.length < 4) continue;
        const a = list[0].a, ux = Math.cos(a), uy = Math.sin(a);
        const proj = (p) => p.x * ux + p.y * uy;
        const items = list.map(g => {
            const t0 = proj(g.p0), t1 = proj(g.p1);
            return { ...g, lo: Math.min(t0, t1), hi: Math.max(t0, t1), start: t0 <= t1 ? g.p0 : g.p1, end: t0 <= t1 ? g.p1 : g.p0 };
        }).sort((x, y) => x.lo - y.lo);
        const L = items.map(i => i.hi - i.lo);
        const gaps = items.slice(1).map((it, i) => it.lo - items[i].hi);
        const med = (arr) => [...arr].sort((x, y) => x - y)[Math.floor(arr.length / 2)];
        const mL = med(L), mG = med(gaps);
        if (mG <= 0.3 * PT) continue;
        const okL = L.every(v => Math.abs(v - mL) <= 0.2 * mL + 0.5 * PT);
        const okG = gaps.every(v => Math.abs(v - mG) <= 0.2 * mG + 0.5 * PT);
        if (!okL || !okG) continue;
        const lw = items[0].s.ln.w;
        const dPct = Math.round(mL / lw * 100000), spPct = Math.round(mG / lw * 100000);
        // 先頭図形を統合後の 1 本にし、残りを削除する
        const first = items[0];
        let xml = first.s.item.xml;
        const g = lineGeomXml(items[0].start, items[items.length - 1].end);
        xml = xml.replace(XFRM_RE, buildXfrmTag(g.x, g.y, g.cx, g.cy, g.flipH, g.flipV));
        xml = xml.replace(GEOM_RE, '<a:prstGeom prst="line"><a:avLst/></a:prstGeom>');
        xml = xml.replace(/(<a:ln\b[^>]*>[\s\S]*?<\/a:solidFill>)/, `$1<a:custDash><a:ds d="${dPct}" sp="${spPct}"/></a:custDash>`);
        first.s.item.xml = xml;
        for (const it of items.slice(1)) removed.add(it.idx);
        stats.mergedDashSegments += items.length;
        stats.mergedDashLines += 1;
    }
    return removed;
}


// 三角形（矢尻）が線の端点 P（隣の点 Q から外向き）に付いているかを判定し、距離を返す。付いていなければ null
function arrowEndDistance(g, P, Q) {
    const dl = dist(P, Q);
    if (dl < 1) return null;
    const out = { x: (P.x - Q.x) / dl, y: (P.y - Q.y) / dl };
    const d = Math.min(dist(P, g.tip), dist(P, g.baseMid));
    if (d > Math.max(3 * PT, 0.6 * g.axisLen)) return null;
    if (out.x * g.dir.x + out.y * g.dir.y < 0.9) return null;
    return d;
}

// 1 本のカギ線が複数の直線に分解されているもの（端点どうしが接する直線の連鎖）を 1 本の折れ線へ統合する。
// 誤爆を避けるため、連鎖の端に矢尻（同色の三角形）が付いているものだけを対象にする。
function mergeSegmentChains(tree, stats) {
    let infos = tree.items.map(it => parseShape(it));
    const segs = [];
    infos.forEach((s, i) => {
        if (!s || s.item.tag !== 'sp' || s.kind !== 'line' || s.hasText || !s.lnVisible) return;
        if (s.ln.head || s.ln.tail) return;
        const p = linePoints(s);
        if (!p || dist(p[0], p[1]) < 1) return;
        segs.push({ i, s, p0: p[0], p1: p[1] });
    });
    if (segs.length < 2) return;
    const tris = infos.map(s => (s && s.kind === 'triangle' && s.fill.kind === 'solid' && !s.lnVisible && !s.hasText) ? triangleGeom(s) && { g: triangleGeom(s), color: s.fill.color } : null).filter(Boolean);
    if (!tris.length) return;
    const tol = 1.5 * PT;
    const same = (a, b) => a.s.ln.color === b.s.ln.color && a.s.ln.w === b.s.ln.w && (a.s.ln.dash || '') === (b.s.ln.dash || '') && !a.s.ln.custDash && !b.s.ln.custDash;
    // 端点（seg の 0=始点, 1=終点）ごとに、接する相手端点を集める
    const ends = [];
    segs.forEach((g, gi) => { ends.push({ gi, e: 0, p: g.p0 }); ends.push({ gi, e: 1, p: g.p1 }); });
    const partner = new Map(); // `${gi}:${e}` → {gi,e} 。ちょうど 1 つの相手と接し、かつ同種の線のときだけ
    for (const a of ends) {
        const near = ends.filter(b => b.gi !== a.gi && dist(a.p, b.p) <= tol);
        if (near.length === 1 && same(segs[a.gi], segs[near[0].gi])) partner.set(`${a.gi}:${a.e}`, near[0]);
    }
    const used = new Set();
    const drop = new Set();
    for (let gi = 0; gi < segs.length; gi++) {
        if (used.has(gi)) continue;
        // gi から連鎖を両方向にたどる（互いが相手になっている端点のみ）
        const chain = [{ gi, flip: false }];
        const seen = new Set([gi]);
        const walk = (startGi, startE, push) => {
            let cur = { gi: startGi, e: startE };
            for (;;) {
                const nx = partner.get(`${cur.gi}:${cur.e}`);
                if (!nx || seen.has(nx.gi)) return;
                const back = partner.get(`${nx.gi}:${nx.e}`);
                if (!back || back.gi !== cur.gi || back.e !== cur.e) return;
                seen.add(nx.gi);
                // 次の線は nx.e 側から入り、反対側（1 - nx.e）へ抜ける
                push({ gi: nx.gi, flip: nx.e === 1 });
                cur = { gi: nx.gi, e: 1 - nx.e };
            }
        };
        walk(gi, 1, c => chain.push(c));
        const head = [];
        walk(gi, 0, c => head.unshift({ gi: c.gi, flip: !c.flip }));
        const all = [...head, ...chain];
        if (all.length < 2) continue;
        // 点列を組み立てる
        const pts = [];
        all.forEach((c, k) => {
            const sg = segs[c.gi];
            const a = c.flip ? sg.p1 : sg.p0, b = c.flip ? sg.p0 : sg.p1;
            if (k === 0) pts.push({ ...a });
            pts.push({ ...b });
        });
        const color = segs[all[0].gi].s.ln.color;
        // 端に同色の矢尻が付いているか
        const endAt = (P, Q) => tris.some(t => t.color === color && arrowEndDistance(t.g, P, Q) != null);
        const ok = endAt(pts[0], pts[1]) || endAt(pts[pts.length - 1], pts[pts.length - 2]);
        if (!ok) continue;
        // 先頭の線を折れ線に作り替え、残りを削除する
        const first = segs[all[0].gi];
        let xml = first.s.item.xml;
        xml = rebuildPolylineXml(xml, pts);
        first.s.item.xml = xml;
        for (const c of all) { used.add(c.gi); if (c.gi !== all[0].gi) drop.add(segs[c.gi].i); }
        stats.chainsMerged++;
        stats.chainSegments += all.length;
    }
    if (drop.size) tree.items = tree.items.filter((_, i) => !drop.has(i));
}

// ---------------------------------------------------------
// スライド 1 枚の再構築
// ---------------------------------------------------------
export function refineSlide(xml, { slideW, slideH }) {
    const stats = {
        negativeFixed: 0, nativeRects: 0, mergedDashLines: 0, mergedDashSegments: 0,
        arrowheadsMerged: 0, arrowheadsLeft: 0, connectorsCreated: 0, glueEnds: 0,
        chainsMerged: 0, chainSegments: 0,
        dashToPreset: 0, dashKept: 0, fontsPaired: 0, fontsReplaced: 0, textboxesWidened: 0,
    };
    const neg = fixNegativeSizes(xml);
    xml = neg.xml; stats.negativeFixed = neg.fixed;

    const tree = splitSpTree(xml);
    if (tree) {
        let infos = tree.items.map(it => parseShape(it));

        // 2. 軸平行四角形の custGeom → native rect
        infos.forEach((s) => {
            if (!s || s.item.tag !== 'sp' || s.kind !== 'rectlike' || !s.cust) return;
            if (s.xf.cx <= 0 || s.xf.cy <= 0) return;
            s.item.xml = s.item.xml.replace(/<a:custGeom>[\s\S]*?<\/a:custGeom>/, '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>');
            stats.nativeRects++;
        });
        infos = tree.items.map(it => parseShape(it));

        // 3. 短線の連なりの破線を 1 本へ
        const removed = mergeDashedSegments(infos, stats);
        if (removed.size) {
            tree.items = tree.items.filter((_, i) => !removed.has(i));
            infos = tree.items.map(it => parseShape(it));
        }

        // 3b. 分解されたカギ線（端点どうしが接する直線の連鎖）を 1 本の折れ線へ
        mergeSegmentChains(tree, stats);
        infos = tree.items.map(it => parseShape(it));

        // 4. 三角矢尻を線の端点へ統合
        const lines = infos.map((s, i) => ({ s, i })).filter(({ s }) => s && s.item.tag === 'sp' && ['line', 'polyline'].includes(s.kind) && !s.hasText && s.lnVisible && !s.fill.color);
        const tris = infos.map((s, i) => ({ s, i })).filter(({ s }) => s && s.kind === 'triangle' && s.fill.kind === 'solid' && !s.lnVisible && !s.hasText);
        const lineState = new Map(); // i → {pts, head, tail}
        for (const { s, i } of lines) lineState.set(i, { pts: linePoints(s).map(p => ({ ...p })), head: null, tail: null, info: s });
        const drop = new Set();
        for (const { s: tri, i: ti } of tris) {
            const g = triangleGeom(tri);
            if (!g) { stats.arrowheadsLeft++; continue; }
            let best = null;
            for (const [li, st] of lineState) {
                if (st.info.ln.color !== tri.fill.color) continue;
                for (const which of ['start', 'end']) {
                    const P = which === 'start' ? st.pts[0] : st.pts[st.pts.length - 1];
                    const Q = which === 'start' ? st.pts[1] : st.pts[st.pts.length - 2];
                    const d = arrowEndDistance(g, P, Q);
                    if (d == null) continue;
                    if (st[which === 'start' ? 'head' : 'tail']) continue;
                    if (!best || d < best.d) best = { li, which, d };
                }
            }
            if (!best) { stats.arrowheadsLeft++; continue; }
            const st = lineState.get(best.li);
            const lw = Math.max(st.info.ln.w, 1);
            const end = { type: 'triangle', w: arrowSize(g.baseW / lw), len: arrowSize(g.axisLen / lw) };
            // 端点は矢尻の先端へ延ばす。ただし線の軸上へ射影し、ほぼ直線なら曲がりを作らない
            const P = best.which === 'start' ? st.pts[0] : st.pts[st.pts.length - 1];
            const Q = best.which === 'start' ? st.pts[1] : st.pts[st.pts.length - 2];
            const dl = dist(P, Q);
            const ux = (P.x - Q.x) / dl, uy = (P.y - Q.y) / dl;
            const tproj = (g.tip.x - P.x) * ux + (g.tip.y - P.y) * uy;
            const cand = { x: P.x + ux * tproj, y: P.y + uy * tproj };
            const newEnd = dist(cand, g.tip) <= 1.5 * PT ? cand : { ...g.tip };
            if (best.which === 'start') { st.pts[0] = newEnd; st.head = end; }
            else { st.pts[st.pts.length - 1] = newEnd; st.tail = end; }
            drop.add(ti);
            stats.arrowheadsMerged++;
        }

        // 5. 直線 → コネクタ（接着）／折れ線は点列と矢印のみ更新
        const nodes = collectNodes(infos, slideW, slideH);
        const tol = GLUE_TOL_PT * PT;
        const siteFor = (P) => {
            let best = null;
            for (const n of nodes) {
                if (n.kind !== 'rect' || !['rect', 'roundRect'].includes(n.prst)) continue;
                for (const site of nodeSites(n)) {
                    const d = dist(P, site);
                    if (d <= tol && (!best || d < best.d)) best = { d, id: n.id, idx: site.idx, x: site.x, y: site.y };
                }
            }
            return best;
        };
        for (const [li, st] of lineState) {
            const info = st.info;
            let xmlShape = info.item.xml;
            const changed = st.head || st.tail || info.kind === 'line';
            if (!changed) continue;
            if (info.kind === 'line' && info.item.tag === 'sp') {
                let [p0, p1] = st.pts;
                // 矢印は必ず終点側に揃える（始点側にだけ矢尻がある場合は線を反転）
                let head = st.head, tail = st.tail;
                if (head && !tail) { [p0, p1] = [p1, p0]; tail = head; head = null; }
                let a = siteFor(p0), b = siteFor(p1);
                // 矢印の無い線は「異なる 2 ノードの接続点に両端が載る」場合だけ接着する
                // （表の罫線・引出線を図形に結び付けると、図形を動かしたとき線が歪むため）
                if (!head && !tail && !(a && b && a.id !== b.id)) { a = null; b = null; }
                if (a) { p0 = { x: a.x, y: a.y }; stats.glueEnds++; }
                if (b) { p1 = { x: b.x, y: b.y }; stats.glueEnds++; }
                const tmp = { ...info, id: info.id };
                let cxn = toConnector(tmp, p0, p1, a, b);
                if (head || tail) cxn = setLineEnds(cxn, { head, tail });
                info.item.tag = 'cxnSp';
                info.item.xml = cxn;
                stats.connectorsCreated++;
            } else {
                // 折れ線（または prst 以外の 2 点線）: 先端が動いた場合のみ再構築
                xmlShape = rebuildPolylineXml(xmlShape, st.pts);
                if (st.head || st.tail) xmlShape = setLineEnds(xmlShape, { head: st.head, tail: st.tail });
                info.item.xml = xmlShape;
            }
        }
        if (drop.size) tree.items = tree.items.filter((_, i) => !drop.has(i));
        xml = joinSpTree(tree);
    }

    // 6. 破線の線スタイル化
    const dash = normalizeDashStyles(xml);
    xml = dash.xml; stats.dashToPreset = dash.toPreset; stats.dashKept = dash.kept;

    // 7. フォント（eastAsia / latin を同一フェイスに）
    const fonts = normalizeFonts(xml);
    xml = fonts.xml; stats.fontsPaired = fonts.paired; stats.fontsReplaced = fonts.replaced;

    // 8. 透過 TextBox の右マージン
    const w = widenTextBoxes(xml, slideW);
    xml = w.xml; stats.textboxesWidened = w.widened;

    return { xml, stats };
}

// ---------------------------------------------------------
// 8. TextBox の右マージン
// ---------------------------------------------------------
export function widenTextBoxes(xml, slideW) {
    let widened = 0;
    const out = xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (sp) => {
        if (!/<p:cNvSpPr[^>]*\btxBox="1"/.test(sp)) return sp;
        if (!/<a:bodyPr[^>]*\bwrap="none"/.test(sp)) return sp;
        if (/<a:bodyPr[^>]*\bvert="(?!horz)[^"]+"/.test(sp)) return sp;
        if (/\balgn="(ctr|r|just|dist|justLow|thaiDist)"/.test(sp)) return sp;
        const xf = parseXfrm(sp);
        if (!xf || xf.rot) return sp;
        const room = slideW - (xf.x + xf.cx);
        const add = Math.min(TEXTBOX_RIGHT_MARGIN_PT * PT, room);
        if (add < 5 * PT) return sp;
        widened++;
        return sp.replace(XFRM_RE, buildXfrmTag(xf.x, xf.y, xf.cx + add, xf.cy, xf.flipH, xf.flipV, xf.rot));
    });
    return { xml: out, widened };
}

// ---------------------------------------------------------
// PPTX 全体への適用
// ---------------------------------------------------------
export function readSlideSize(presXml) {
    const m = /<p:sldSz\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/.exec(presXml) || /<p:sldSz\b[^>]*\bcy="(\d+)"[^>]*\bcx="(\d+)"/.exec(presXml);
    if (!m) return { slideW: 12192000, slideH: 6858000 };
    return /cx="(\d+)"[^>]*cy=/.test(m[0]) ? { slideW: parseInt(m[1], 10), slideH: parseInt(m[2], 10) } : { slideW: parseInt(m[2], 10), slideH: parseInt(m[1], 10) };
}

const slideNum = (p) => parseInt((p.match(/slide(\d+)\.xml$/) || [0, 0])[1], 10);

export async function refinePptx(pptxPath, { fontMode = 'auto' } = {}) {
    const zip = await JSZip.loadAsync(fs.readFileSync(pptxPath));
    const pres = zip.file('ppt/presentation.xml');
    const size = readSlideSize(pres ? await pres.async('string') : '');
    const slideFiles = Object.keys(zip.files).filter(p => /^ppt\/slides\/slide\d+\.xml$/.test(p)).sort((a, b) => slideNum(a) - slideNum(b));
    const total = {
        slides: slideFiles.length, negativeFixed: 0, nativeRects: 0, mergedDashLines: 0, mergedDashSegments: 0,
        arrowheadsMerged: 0, arrowheadsLeft: 0, connectorsCreated: 0, glueEnds: 0,
        chainsMerged: 0, chainSegments: 0,
        dashToPreset: 0, dashKept: 0, fontsPaired: 0, fontsReplaced: 0, textboxesWidened: 0,
    };
    // 全スライドを先に計算し、1 枚でも例外なら何も書き換えない
    const results = [];
    for (const f of slideFiles) {
        const xml = await zip.file(f).async('string');
        const r = refineSlide(xml, size);
        results.push([f, r]);
    }
    for (const [f, r] of results) {
        zip.file(f, r.xml);
        for (const k of Object.keys(r.stats)) total[k] += r.stats[k];
    }
    fs.writeFileSync(pptxPath, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
    return total;
}

// =========================================================
// 機械検品（V1.5 仕上げの①〜⑥）
// 出力 PPTX を読み戻して数える。数えていない項目に「OK」は付けない。
// =========================================================
const DASH_PRESET_KIND = {
    sysDot: 'dot', dot: 'dot',
    dash: 'dash', lgDash: 'dash', sysDash: 'dash', dashDot: 'dash', lgDashDot: 'dash', lgDashDotDot: 'dash', sysDashDot: 'dash', sysDashDotDot: 'dash',
};

export async function auditPptx(pptxPath) {
    const zip = await JSZip.loadAsync(fs.readFileSync(pptxPath));
    const pres = zip.file('ppt/presentation.xml');
    const { slideW, slideH } = readSlideSize(pres ? await pres.async('string') : '');
    const slideFiles = Object.keys(zip.files).filter(p => /^ppt\/slides\/slide\d+\.xml$/.test(p)).sort((a, b) => slideNum(a) - slideNum(b));

    const report = {
        slideSize: { wIn: +(slideW / 914400).toFixed(3), hIn: +(slideH / 914400).toFixed(3) },
        slides: [],
        connectors: { total: 0, candidates: 0, glued: 0, bothLanded: 0, oneLanded: 0, noneLanded: 0, decorative: 0, polylines: 0, edges: [] },
        lineStyles: { solid: 0, dash: 0, dot: 0, custom: 0 },
        negativeSize: 0,
        bold: { total: 0, ruleOk: 0, review: 0 },
        fonts: { faces: {}, nonStandard: [], mismatch: 0, missingPair: 0, inherited: 0 },
        bounds: { outOfSlide: [], textOverflow: [] },
        pictures: [],
        raw: { arrowheadLeftovers: 0 },
    };

    for (const f of slideFiles) {
        const n = slideNum(f);
        const xml = await zip.file(f).async('string');
        const tree = splitSpTree(xml);
        const slide = { no: n, shapes: 0, textBoxes: 0, connectors: 0, pictures: 0, groups: 0, textRuns: 0 };

        // ③ 負サイズ（grpSp 内も含め全 xfrm を数える）
        for (const m of xml.matchAll(/<a:ext cx="(-?\d+)" cy="(-?\d+)"\s*\/>/g)) {
            if (parseInt(m[1], 10) < 0 || parseInt(m[2], 10) < 0) report.negativeSize++;
        }

        // ④ 太字 / ⑤ フォント
        for (const m of xml.matchAll(/<a:r>\s*<a:rPr\b([^>]*?)(?:\/>|>([\s\S]*?)<\/a:rPr>)\s*<a:t>([\s\S]*?)<\/a:t>/g)) {
            slide.textRuns++;
            const attrs = m[1] || '', inner = m[2] || '', text = decodeEntities(m[3]);
            if (text.trim() === '') continue;
            const szPt = parseInt((attrs.match(/\bsz="(\d+)"/) || [0, 0])[1], 10) / 100;
            if (/\bb="1"/.test(attrs)) {
                report.bold.total++;
                // 太字にしてよい: 24pt 以上 / 英字ラベル（CJK を含まない）
                if (szPt >= 24 || !hasCjk(text)) report.bold.ruleOk++; else report.bold.review++;
            }
            const lat = (/<a:latin typeface="([^"]*)"/.exec(inner) || [])[1];
            const ea = (/<a:ea typeface="([^"]*)"/.exec(inner) || [])[1];
            for (const face of [lat, ea]) if (face) report.fonts.faces[decodeEntities(face)] = (report.fonts.faces[decodeEntities(face)] || 0) + 1;
            if (!lat && !ea) report.fonts.inherited++;
            else if (!lat || !ea) report.fonts.missingPair++;
            else if (lat !== ea) report.fonts.mismatch++;
        }

        if (!tree) { report.slides.push(slide); continue; }
        const infos = tree.items.map(it => parseShape(it));
        tree.items.forEach(it => {
            if (it.tag === 'pic') {
                slide.pictures++;
                const x = parseXfrm(it.xml);
                if (x) report.pictures.push({ slide: n, wPt: +(x.cx / PT).toFixed(0), hPt: +(x.cy / PT).toFixed(0), coversSlide: x.cx * x.cy >= 0.6 * slideW * slideH });
            } else if (it.tag === 'grpSp') slide.groups++;
        });
        slide.shapes = infos.filter(Boolean).length;
        slide.textBoxes = infos.filter(s => s && s.isTxBox).length;

        // ⑥ スライド内に収まっているか / 文字幅
        for (const s of infos) {
            if (!s) continue;
            const { x, y, cx, cy } = s.xf;
            const tol = 1 * PT;
            if (x < -tol || y < -tol || x + cx > slideW + tol || y + cy > slideH + tol) {
                report.bounds.outOfSlide.push({ slide: n, id: s.id, kind: s.kind });
            }
            if (s.hasText && /wrap="none"/.test(s.item.xml)) {
                for (const para of s.item.xml.match(/<a:p>[\s\S]*?<\/a:p>/g) || []) {
                    let w = 0;
                    for (const r of para.match(/<a:r>[\s\S]*?<\/a:r>/g) || []) {
                        const sz = parseInt((r.match(/\bsz="(\d+)"/) || [0, 1800])[1], 10) / 100;
                        const t = decodeEntities((/<a:t>([\s\S]*?)<\/a:t>/.exec(r) || [0, ''])[1]);
                        const spc = parseInt((r.match(/\bspc="(-?\d+)"/) || [0, 0])[1], 10) / 100;
                        w += estimateTextWidthPt(t, sz) + spc * [...t].length;
                    }
                    const lIns = parseInt((s.item.xml.match(/\blIns="(-?\d+)"/) || [0, 91440])[1], 10);
                    const rIns = parseInt((s.item.xml.match(/\brIns="(-?\d+)"/) || [0, 91440])[1], 10);
                    const inner = (cx - lIns - rIns) / PT;
                    if (w > inner + 0.5) { report.bounds.textOverflow.push({ slide: n, id: s.id, estPt: +w.toFixed(1), boxPt: +inner.toFixed(1) }); break; }
                }
            }
        }

        // ① 接続線 / ② 線種
        const nodes = collectNodes(infos, slideW, slideH);
        const tolEmu = LANDING_TOL_PT * PT;
        const landed = (P) => {
            for (const nd of nodes) if (nodeEdgeDistance(nd, P) <= tolEmu) return nd;
            return null;
        };
        const tris = infos.filter(s => s && s.kind === 'triangle' && s.fill.kind === 'solid' && !s.lnVisible && triangleGeom(s));
        report.raw.arrowheadLeftovers += tris.length;
        for (const s of infos) {
            if (!s || !['line', 'polyline', 'connector'].includes(s.kind)) continue;
            if (s.hasText || (!s.lnVisible)) continue;
            const pts = linePoints(s);
            if (!pts) continue;
            // 線種
            if (s.ln.custDash) report.lineStyles.custom++;
            else if (s.ln.dash && DASH_PRESET_KIND[s.ln.dash]) report.lineStyles[DASH_PRESET_KIND[s.ln.dash]]++;
            else report.lineStyles.solid++;

            const isCxn = s.item.tag === 'cxnSp';
            report.connectors.total++;
            slide.connectors++;
            if (s.kind === 'polyline') report.connectors.polylines++;
            const A = pts[0], B = pts[pts.length - 1];
            const la = landed(A), lb = landed(B);
            const st = isCxn ? (/<a:stCxn id="(\d+)" idx="(\d+)"/.exec(s.item.xml) || null) : null;
            const en = isCxn ? (/<a:endCxn id="(\d+)" idx="(\d+)"/.exec(s.item.xml) || null) : null;
            const arrow = (s.ln.head && s.ln.head.type !== 'none' ? 1 : 0) + (s.ln.tail && s.ln.tail.type !== 'none' ? 2 : 0);
            // 罫線（表の枠線・同一ノードの輪郭に沿う線）は接続線にしない。
            // 接続線 = 矢印付き / 接着済み / 異なる 2 ノードの辺に両端が着地し、かつ輪郭に沿っていない線
            const mids = [];
            for (let k = 0; k + 1 < pts.length; k++) mids.push({ x: (pts[k].x + pts[k + 1].x) / 2, y: (pts[k].y + pts[k + 1].y) / 2 });
            const alongEdge = mids.every(m => [la, lb].some(nd => nd && nodeEdgeDistance(nd, m) <= tolEmu));
            const bridges = !!(la && lb && la !== lb && !alongEdge);
            const whiteRule = !arrow && (s.ln.color || '') === 'ffffff'; // 白い区切り線は罫線扱い
            const isCand = !!(arrow || st || en || (bridges && !whiteRule)) && dist(A, B) >= 6 * PT;
            if (!isCand) { report.connectors.decorative++; continue; }
            report.connectors.candidates++;
            if (st) report.connectors.glued++;
            if (en) report.connectors.glued++;
            const nl = (la ? 1 : 0) + (lb ? 1 : 0);
            if (nl === 2) report.connectors.bothLanded++; else if (nl === 1) report.connectors.oneLanded++; else report.connectors.noneLanded++;
            const dashName = s.ln.custDash ? `custDash(${(s.ln.custDash.d / 1000).toFixed(0)}%/${(s.ln.custDash.sp / 1000).toFixed(0)}%)` : (s.ln.dash || 'solid');
            report.connectors.edges.push({
                no: report.connectors.edges.length + 1, slide: n,
                from: st ? `#${st[1]}` : la ? `#${la.id}(辺上・非接着)` : '自由端',
                to: en ? `#${en[1]}` : lb ? `#${lb.id}(辺上・非接着)` : '自由端',
                color: (s.ln.color || '').toUpperCase(),
                style: dashName,
                arrow: arrow === 0 ? 'なし' : arrow === 1 ? '始点側' : arrow === 2 ? '終点側' : '両端',
                glued: !!(st && en) ? '両端接着' : (st || en) ? '片端接着' : '非接着',
                landed: nl,
                kind: s.kind === 'polyline' ? '折れ線' : '直線',
            });
        }
        report.slides.push(slide);
    }

    // ⑤ 非標準フォント
    report.fonts.nonStandard = Object.keys(report.fonts.faces).filter(n => !isStandardFont(decodeEntities(n)));
    return report;
}

// =========================================================
// 納品メモ（V1.5 の 4 見出し）
// =========================================================
export function buildDeliveryNote({ fileName = '', version = '', refine = null, audit, route = '' }) {
    const a = audit, c = a.connectors, L = [];
    const nSlides = a.slides.length;
    const sum = (k) => a.slides.reduce((s, x) => s + x[k], 0);
    L.push(`# 納品メモ — ${fileName || '変換結果'}`);
    L.push(`PPTX Converter ${version} ／ ${nSlides} スライド ／ ${a.slideSize.wIn}×${a.slideSize.hIn} in${route ? ` ／ 入力経路: ${route}` : ''}`);
    L.push('');
    L.push('## ① 図版の処理方法');
    const covering = a.pictures.filter(p => p.coversSlide);
    L.push(`- 図形・テキスト: ネイティブ図形 ${sum('shapes')} 個（うち TextBox ${sum('textBoxes')}）／ グループ ${sum('groups')}`);
    L.push(`- 画像: PNG 切り出し ${sum('pictures')} 枚（写真・複雑な質感・人物・ロゴは図形化しない方針）${covering.length ? `。スライド面積 60% 以上の全面画像が ${covering.length} 枚残存` : ''}`);
    if (refine) {
        L.push(`- 四角形: 自由曲線から native な rect へ復元 ${refine.nativeRects} 個`);
        L.push(`- 接続線: 直線コネクタ化 ${refine.connectorsCreated} 本／矢尻を線の端に統合 ${refine.arrowheadsMerged} 個／分解されたカギ線の統合 ${refine.chainsMerged || 0} 本（${refine.chainSegments || 0} 線分→1 本）／自作破線(短線の連なり)の統合 ${refine.mergedDashLines} 本`);
    }
    L.push(`- エッジ表: ${c.edges.length} 行 ＝ 出力中の接続線候補 ${c.candidates} 本（装飾線・罫線 ${c.decorative} 本は除外）。照合結果: ${c.edges.length === c.candidates ? '一致' : '不一致'}`);
    if (c.edges.length) {
        L.push('');
        L.push('| No. | 頁 | 始点 | 終点 | 色 | 線種 | 矢印 | 接着 |');
        L.push('|---|---|---|---|---|---|---|---|');
        for (const e of c.edges.slice(0, 60)) L.push(`| ${e.no} | ${e.slide} | ${e.from} | ${e.to} | ${e.color} | ${e.style} | ${e.arrow} | ${e.glued} |`);
        if (c.edges.length > 60) L.push(`| … | | 他 ${c.edges.length - 60} 行は JSON（?format=json）参照 | | | | | |`);
    }
    L.push('');
    L.push('## ② 近似・目測・非接着・要確認の箇所');
    const notes = [];
    const notGlued = c.edges.filter(e => e.glued === '非接着').length;
    if (notGlued) notes.push(`非接着の接続線 ${notGlued} 本（端点が rect の辺の中点にない／ノードが rect 以外／折れ線のため）。図形を動かしても追従しません`);
    if (c.polylines) notes.push(`折れ線（カギ線）は 1 オブジェクトのまま維持したが、コネクタではなく自由曲線（非接着）: ${c.polylines} 本`);
    if (a.raw.arrowheadLeftovers) notes.push(`矢尻が別シェイプのまま残っている三角形 ${a.raw.arrowheadLeftovers} 個（線の端点と一致しない）`);
    if (c.noneLanded) notes.push(`両端とも辺上に着地していない接続線 ${c.noneLanded} 本（ノードが rect/楕円以外、または自由端）`);
    if (a.lineStyles.custom) notes.push(`カスタム破線 ${a.lineStyles.custom} 本（dash/sysDash/sysDot に近似できず custDash のまま）`);
    if (a.bold.review) notes.push(`太字の要確認 ${a.bold.review} 箇所（24pt 未満で日本語を含む＝原画で太字に見えるか目視確認）`);
    if (a.fonts.nonStandard.length) notes.push(`標準外フォント: ${a.fonts.nonStandard.join(', ')}`);
    if (a.bounds.outOfSlide.length) notes.push(`スライド外にはみ出す図形 ${a.bounds.outOfSlide.length} 個（全面背景のブリードの可能性。要確認）`);
    if (a.bounds.textOverflow.length) notes.push(`文字幅の推定が箱幅を超える TextBox ${a.bounds.textOverflow.length} 個`);
    notes.push('文字・位置・色は PDF インポート値に基づく（目測なし）。原画との並置確認は PowerPoint で 1 回だけ目視してください（本サーバーでは画像化して比較していません）');
    notes.forEach(n => L.push(`- ${n}`));
    L.push('');
    L.push('## ③ 機械検品の実測値');
    L.push(`- ① 接続線の着地: 候補 ${c.candidates} 本 ／ 両端が辺±${LANDING_TOL_PT}pt に着地 ${c.bothLanded} ／ 片端のみ ${c.oneLanded} ／ なし ${c.noneLanded} ／ 接着端 ${c.glued}（接着先 ID は上表）`);
    L.push(`- ② 線種（接続線・装飾線を含む全${c.total}本）: 実線 ${a.lineStyles.solid} ／ 破線 ${a.lineStyles.dash} ／ 点線 ${a.lineStyles.dot} ／ カスタム破線 ${a.lineStyles.custom}（原画の本数との一致は未照合）`);
    L.push(`- ③ 負サイズの図形: ${a.negativeSize} 個${a.negativeSize === 0 ? '' : '（要修正）'}`);
    L.push(`- ④ 太字: ${a.bold.total} 箇所（ルール適合 ${a.bold.ruleOk} ／ 要確認 ${a.bold.review}）。原画の太字数との一致は未照合`);
    L.push(`- ⑤ フォント: eastAsia と latin が不一致 ${a.fonts.mismatch} ／ 片方のみ ${a.fonts.missingPair} ／ 標準外 ${a.fonts.nonStandard.length} 種`);
    L.push(`- ⑥ 範囲・文字幅: スライド外 ${a.bounds.outOfSlide.length} 個 ／ 推定文字幅が箱幅超過 ${a.bounds.textOverflow.length} 個`);
    L.push('');
    L.push('## ④ 次の改善候補');
    const next = [];
    if (notGlued) next.push('非接着の接続線を、PowerPoint 上で図形に接着し直す（または原画のノードが rect 以外なら rect へ置換）');
    if (c.polylines) next.push('カギ線を bentConnector（折れ線コネクタ）へ置き換えると、ノード移動に追従する');
    if (a.bold.review) next.push('太字の要確認箇所を原画と見比べ、本文の太字を外す');
    if (covering.length) next.push('全面画像が残るスライドは「一枚絵の図解」オプション（分解＋文字をテキスト化）で再変換する');
    if (a.bounds.textOverflow.length) next.push('文字幅が超過する TextBox は幅を広げるか文字サイズを 1 段階下げる');
    next.push('配色・文言を差し替える場合は、色は図形の塗り、文字は TextBox を直接編集（画像化された文字は無い前提）');
    next.slice(0, 3).forEach(n => L.push(`- ${n}`));
    return L.join('\n');
}

// CLI: node refine.js <file.pptx> [--json]
if (process.argv[1] && /refine\.js$/.test(process.argv[1])) {
    const file = process.argv[2];
    if (!file) { console.error('usage: node refine.js <file.pptx> [--json]'); process.exit(1); }
    const audit = await auditPptx(file);
    if (process.argv.includes('--json')) console.log(JSON.stringify(audit, null, 2));
    else console.log(buildDeliveryNote({ fileName: file, version: require('./package.json').version, audit }));
}
