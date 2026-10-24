// refine.js（うちた式パワポ変換 V1.5 の後処理・機械検品）のテスト
//   fixtures は実際に LibreOffice の PDF インポートが出力した PPTX（未加工）
//   実行: npm test
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { createRequire } from 'module';
import {
    refinePptx, auditPptx, buildDeliveryNote, fixNegativeSizes, normalizeFonts,
    normalizeDashStyles, nearestDashPreset, widenTextBoxes, estimateTextWidthPt, isStandardFont,
} from '../refine.js';

const require = createRequire(import.meta.url);
const JSZip = require('jszip');
const fx = (n) => path.join(path.dirname(new URL(import.meta.url).pathname), 'fixtures', n);

function tmpCopy(name) {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'refine-'));
    const dst = path.join(dir, name);
    fs.copyFileSync(fx(name), dst);
    return dst;
}

// 簡易の整形式チェック（タグの開閉が対応しているか）
function assertBalanced(xml, label) {
    const stack = [];
    for (const m of xml.matchAll(/<(\/?)([A-Za-z0-9_:.-]+)\b[^>]*?(\/?)>/g)) {
        const [, close, name, self] = m;
        if (name.startsWith('?') || self === '/') continue;
        if (close) assert.equal(stack.pop(), name, `${label}: </${name}> の対応が崩れている`);
        else stack.push(name);
    }
    assert.equal(stack.length, 0, `${label}: 閉じていないタグが残っている ${stack.join(',')}`);
}

async function slideXml(pptx, n = 1) {
    const zip = await JSZip.loadAsync(fs.readFileSync(pptx));
    return zip.file(`ppt/slides/slide${n}.xml`).async('string');
}

test('負のサイズは flip に正規化される', () => {
    const src = '<a:xfrm><a:off x="1000" y="2000"/><a:ext cx="-300" cy="-500"/></a:xfrm>';
    const { xml, fixed } = fixNegativeSizes(src);
    assert.equal(fixed, 1);
    assert.match(xml, /flipH="1"/);
    assert.match(xml, /flipV="1"/);
    assert.match(xml, /<a:off x="700" y="1500"\/><a:ext cx="300" cy="500"\/>/);
});

test('custDash は近い場合だけ線スタイルへ、遠い場合は保持', () => {
    assert.equal(nearestDashPreset(400, 300).name, 'dash');
    assert.equal(nearestDashPreset(300, 100).name, 'sysDash');
    assert.equal(nearestDashPreset(100, 100).name, 'sysDot');
    const near = normalizeDashStyles('<a:custDash><a:ds d="300000" sp="100000"/></a:custDash>');
    assert.equal(near.toPreset, 1);
    assert.match(near.xml, /<a:prstDash val="sysDash"\/>/);
    const far = normalizeDashStyles('<a:custDash><a:ds d="74000" sp="150000"/></a:custDash>');
    assert.equal(far.kept, 1);
    assert.match(far.xml, /custDash/);
});

test('latin と ea が同じフェイスになり、非標準フォントは標準へ置換される', () => {
    const run = (faces, text) => `<a:r><a:rPr lang="ja-JP" sz="1400">${faces}</a:rPr><a:t>${text}</a:t></a:r>`;
    // ea だけ・latin だけ
    let r = normalizeFonts(run('<a:ea typeface="Meiryo"/>', 'あ'));
    assert.match(r.xml, /<a:latin typeface="Meiryo"\/><a:ea typeface="Meiryo"\/>/);
    // 非標準（Noto 系は標準外）
    r = normalizeFonts(run('<a:latin typeface="Noto Sans CJK JP"/><a:ea typeface="Noto Sans CJK JP"/>', '日本語'));
    assert.match(r.xml, /<a:latin typeface="Yu Gothic"\/><a:ea typeface="Yu Gothic"\/>/);
    assert.equal(r.replaced, 1);
    r = normalizeFonts(run('<a:latin typeface="DejaVu Sans"/><a:ea typeface="DejaVu Sans"/>', 'ABC'));
    assert.match(r.xml, /<a:latin typeface="Arial"\/><a:ea typeface="Arial"\/>/);
    // 自己終了 rPr が次の run を巻き込まない
    const two = '<a:r><a:rPr b="1"/><a:t>X</a:t></a:r>' + run('<a:latin typeface="Arial"/>', 'Y');
    r = normalizeFonts(two);
    assert.match(r.xml, /<a:r><a:rPr b="1"\/><a:t>X<\/a:t><\/a:r>/);
    assertBalanced(`<a:p>${r.xml}</a:p>`, 'fonts');
    // テーマ参照は触らない
    const theme = run('<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/>', 'あ');
    assert.equal(normalizeFonts(theme).xml, theme);
    assert.ok(isStandardFont('Yu Gothic') && isStandardFont('ＭＳ ゴシック') && !isStandardFont('Noto Sans CJK JP'));
});

test('透過 TextBox の右マージンを広げる（中央・右揃えは対象外、スライド端でクランプ）', () => {
    const sp = (x, algn = '') => `<p:sp><p:nvSpPr><p:cNvPr id="2" name=""/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>` +
        `<a:xfrm><a:off x="${x}" y="0"/><a:ext cx="1000000" cy="200000"/></a:xfrm></p:spPr>` +
        `<p:txBody><a:bodyPr wrap="none"/><a:p>${algn ? `<a:pPr algn="${algn}"/>` : ''}<a:r><a:t>x</a:t></a:r></a:p></p:txBody></p:sp>`;
    const W = 12192000;
    assert.match(widenTextBoxes(sp(100000), W).xml, /cx="1381000"/);
    assert.equal(widenTextBoxes(sp(100000, 'ctr'), W).widened, 0);
    assert.equal(widenTextBoxes(sp(100000, 'r'), W).widened, 0);
    // 右端ぎりぎり: 余白が無ければ広げない
    assert.equal(widenTextBoxes(sp(W - 1000000), W).widened, 0);
});

test('文字幅の推定は 全角=サイズ×1.0 / 半角=×0.56', () => {
    assert.equal(estimateTextWidthPt('あい', 10), 20);
    assert.ok(Math.abs(estimateTextWidthPt('AB', 10) - 11.2) < 1e-9);
});

test('実機 fixture(1): 矢尻を統合しコネクタ化して両端を接着する', async () => {
    const f = tmpCopy('lo-diagram.pptx');
    const before = await auditPptx(f);
    assert.equal(before.connectors.glued, 0);
    assert.equal(before.connectors.candidates, 3);   // 矢尻が別シェイプでも、辺上に着地した線は候補に数える

    const stats = await refinePptx(f);
    assert.equal(stats.nativeRects, 4);
    assert.equal(stats.arrowheadsMerged, 3);
    assert.equal(stats.connectorsCreated, 3);
    assert.equal(stats.glueEnds, 6);

    const xml = await slideXml(f);
    assertBalanced(xml, 'diagram');
    assert.equal((xml.match(/<p:cxnSp>/g) || []).length, 3);
    assert.equal((xml.match(/<a:tailEnd type="triangle"/g) || []).length, 3);
    assert.equal((xml.match(/<a:stCxn /g) || []).length, 3);
    assert.equal((xml.match(/<a:endCxn /g) || []).length, 3);

    const after = await auditPptx(f);
    assert.equal(after.connectors.candidates, 3);
    assert.equal(after.connectors.glued, 6);
    assert.equal(after.connectors.bothLanded, 3);
    assert.equal(after.raw.arrowheadLeftovers, 0);
    assert.equal(after.negativeSize, 0);
    assert.equal(after.fonts.mismatch, 0);
    assert.equal(after.fonts.missingPair, 0);
    assert.deepEqual(after.fonts.nonStandard, []);
    // 太字: KEY TAKEAWAY（英字ラベル）だけが太字でルール適合
    assert.equal(after.bold.total, 1);
    assert.equal(after.bold.review, 0);
    // 線種: 実線 / 破線(dash) / 点線(custDash のまま or sysDot)
    assert.equal(after.lineStyles.solid, 1);
    assert.equal(after.lineStyles.dash + after.lineStyles.custom + after.lineStyles.dot, 2);
    assert.equal(after.connectors.edges.length, after.connectors.candidates);
});

test('実機 fixture(2): 折れ線・斜め線・自作破線（短線の連なり）', async () => {
    const f = tmpCopy('lo-diagram2.pptx');
    const stats = await refinePptx(f);
    assert.equal(stats.mergedDashLines, 1);
    assert.equal(stats.mergedDashSegments, 12);
    assert.equal(stats.arrowheadsMerged, 3);
    assert.equal(stats.connectorsCreated, 3);        // 直線 3 本（接続線 2 + 統合後の破線 1。折れ線は自由曲線のまま）

    const xml = await slideXml(f);
    assertBalanced(xml, 'diagram2');
    const a = await auditPptx(f);
    assert.equal(a.connectors.polylines, 1);
    assert.equal(a.connectors.decorative, 1);        // 統合後の破線 1 本は接続線ではない
    assert.equal(a.lineStyles.custom + a.lineStyles.dash + a.lineStyles.dot, 1);
    assert.equal(a.raw.arrowheadLeftovers, 0);
    assert.equal(a.connectors.edges.filter(e => e.arrow === '終点側').length, 3);   // 矢印は必ず終点側に揃う
    assert.equal(a.connectors.edges.filter(e => e.glued === '両端接着').length, 2);

    // 折れ線の終点は直線のまま（曲がりを作らない）
    const poly = [...xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)].map(m => m[0]).find(s => /<a:tailEnd/.test(s) && /<a:custGeom>/.test(s));
    assert.ok(poly, '折れ線に tailEnd がある');
    const pts = [...poly.matchAll(/<a:pt x="(\d+)" y="(\d+)"\/>/g)].map(m => [+m[1], +m[2]]);
    assert.equal(pts.length, 3);
    assert.equal(pts[1][0], pts[2][0], '最後の線分は垂直のまま');
});

test('実機 fixture(3): 直線に分解されたカギ線＋矢尻は 1 本の折れ線になり、矢印なしの L 字罫線は統合しない', async () => {
    const f = tmpCopy('lo-diagram3.pptx');
    const before = await auditPptx(f);
    assert.equal(before.connectors.total, 5);          // 3 線分 + 罫線 2
    const stats = await refinePptx(f);
    assert.equal(stats.chainsMerged, 1);
    assert.equal(stats.chainSegments, 3);
    assert.equal(stats.arrowheadsMerged, 1);
    const xml = await slideXml(f);
    assertBalanced(xml, 'diagram3');
    const a = await auditPptx(f);
    assert.equal(a.connectors.total, 3);               // 折れ線 1 + 罫線 2（罫線は別オブジェクトのまま）
    assert.equal(a.connectors.polylines, 1);
    assert.equal(a.connectors.candidates, 1);
    assert.equal(a.connectors.decorative, 2);
    assert.equal(a.connectors.edges[0].arrow, '終点側');
    assert.equal(a.connectors.edges[0].landed, 2);     // 両端が辺±2pt に着地
    assert.equal(a.raw.arrowheadLeftovers, 0);
});

test('refine は冪等（2 回かけても XML が壊れず、コネクタ数が増えない）', async () => {
    const f = tmpCopy('lo-diagram.pptx');
    await refinePptx(f);
    const once = await slideXml(f);
    const s2 = await refinePptx(f);
    const twice = await slideXml(f);
    assertBalanced(twice, 'twice');
    assert.equal(s2.connectorsCreated, 0);
    assert.equal((twice.match(/<p:cxnSp>/g) || []).length, (once.match(/<p:cxnSp>/g) || []).length);
});

test('納品メモは 4 見出しを含み、実測値を載せる', async () => {
    const f = tmpCopy('lo-diagram.pptx');
    const r = await refinePptx(f);
    const audit = await auditPptx(f);
    const md = buildDeliveryNote({ fileName: 'x.pptx', version: '9.9.9', refine: r, audit, route: 'pdf' });
    for (const h of ['## ① 図版の処理方法', '## ② 近似・目測・非接着・要確認の箇所', '## ③ 機械検品の実測値', '## ④ 次の改善候補']) {
        assert.ok(md.includes(h), h);
    }
    assert.match(md, /エッジ表: 3 行 ＝ 出力中の接続線候補 3 本/);
    assert.match(md, /照合結果: 一致/);
    assert.match(md, /③ 負サイズの図形: 0 個/);
});
