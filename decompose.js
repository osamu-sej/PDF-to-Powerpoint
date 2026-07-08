// =========================================================
// 一枚絵（フラット化された図解）のパーツ分解
//
// PDF の中に「図解全体が 1 枚のラスタ画像」として埋め込まれて
// いる場合、LibreOffice のインポートではスライド全面を覆う
// 1 枚の画像になってしまい、まったく編集できない。
//
// このモジュールは変換後の PPTX を走査し、スライドの大部分を
// 覆う 1 枚画像を検出したら、画像解析でパーツに分解して
// 「動かせる部品の集まり」として再構築する。
//
//   1. 背景色を推定（最頻色）。写真のように支配的な背景色が
//      無い画像は分解せずそのまま残す（誤爆防止）。
//   2. 背景と異なる画素を前景とみなし、膨張処理で文字や
//      図形のかたまりをまとめてから連結成分に分ける。
//   3. 大きなかたまり（パネル）は内部をさらに再帰的に分解し、
//      「パネルの背景板」と「中身のパーツ」に分ける。
//   4. スライドを「きれいな背景板 + 各パーツ（透過PNG）」で
//      再構築する。見た目は元とほぼ同じまま、各パーツを個別に
//      移動・削除・差し替えできるようになる。
//
// 完全に元の描画命令へ戻す正攻法は存在しないため、これは
// 「多少の劣化を許容して編集可能性を得る」近似的なアプローチ。
// 判定に失敗した場合は必ず元のスライドを無傷で残す。
// =========================================================
import { createRequire } from 'module';
const require = createRequire(import.meta.url);

const fs = require('fs');
const sharp = require('sharp');
const JSZip = require('jszip');

// ---- チューニングパラメータ ----
const PIC_AREA_RATIO_MIN = 0.60;   // スライド面積に対する画像の占有率（これ以上で分解候補）
const MIN_IMAGE_EDGE_PX = 400;     // 分解対象とする画像の最小辺長
const MAX_WORK_EDGE_PX = 3200;     // 解析時の最大辺長（超える場合は縮小して処理）
const COLOR_DIST_T = 70;           // 背景色との距離しきい値 (|Δr|+|Δg|+|Δb|)
const MIN_BG_RATIO = 0.25;         // 背景色の占有率がこれ未満なら「写真」とみなし分解しない
const MIN_COMP_PIXELS = 30;        // これより小さい連結成分はノイズとして背景板に残す
const MIN_COMP_EDGE = 5;           // 連結成分バウンディングボックスの最小辺 (px)
const MAX_PARTS = 160;             // 1画像あたりの最大パーツ数
const MIN_PARTS = 4;               // これ未満しか分解できない場合は元のまま残す
const MAX_DEPTH = 3;               // パネル内再分解の最大深さ

// =========================================================
// 低レベル画像解析（RGBA の生バッファ上で動く）
// =========================================================

// 領域内の最頻色（背景色候補）を求める
// 4bit/チャネルに量子化したヒストグラムで最頻ビンを探し、
// そのビンに属する画素の平均色を背景色として返す
function dominantColor(data, W, rect) {
    const hist = new Uint32Array(4096);
    let total = 0;
    for (let y = rect.y0; y < rect.y1; y++) {
        let i = (y * W + rect.x0) * 4;
        for (let x = rect.x0; x < rect.x1; x++, i += 4) {
            total++;
            if (data[i + 3] < 128) { hist[4095]++; continue; } // 透明は白扱い
            hist[((data[i] >> 4) << 8) | ((data[i + 1] >> 4) << 4) | (data[i + 2] >> 4)]++;
        }
    }
    let best = 0, bestKey = 0;
    for (let k = 0; k < 4096; k++) if (hist[k] > best) { best = hist[k]; bestKey = k; }
    // 最頻ビンに属する画素の平均色を求める
    let sr = 0, sg = 0, sb = 0, n = 0;
    for (let y = rect.y0; y < rect.y1; y++) {
        let i = (y * W + rect.x0) * 4;
        for (let x = rect.x0; x < rect.x1; x++, i += 4) {
            const key = data[i + 3] < 128 ? 4095
                : ((data[i] >> 4) << 8) | ((data[i + 1] >> 4) << 4) | (data[i + 2] >> 4);
            if (key !== bestKey) continue;
            if (data[i + 3] < 128) { sr += 255; sg += 255; sb += 255; }
            else { sr += data[i]; sg += data[i + 1]; sb += data[i + 2]; }
            n++;
        }
    }
    if (n === 0) return null;
    return { r: Math.round(sr / n), g: Math.round(sg / n), b: Math.round(sb / n), ratio: best / total };
}

// 前景マスク（背景色から一定以上離れた画素 = 1）を作る
function buildMask(data, W, rect, bg) {
    const rw = rect.x1 - rect.x0, rh = rect.y1 - rect.y0;
    const mask = new Uint8Array(rw * rh);
    for (let y = 0; y < rh; y++) {
        let i = ((y + rect.y0) * W + rect.x0) * 4;
        let o = y * rw;
        for (let x = 0; x < rw; x++, i += 4, o++) {
            if (data[i + 3] < 128) continue; // 透明は背景
            const d = Math.abs(data[i] - bg.r) + Math.abs(data[i + 1] - bg.g) + Math.abs(data[i + 2] - bg.b);
            if (d > COLOR_DIST_T) mask[o] = 1;
        }
    }
    return mask;
}

// マスクを半径 r で膨張させる（水平→垂直の 2 パスのボックス膨張）
function dilate(mask, rw, rh, r) {
    if (r <= 0) return mask;
    const tmp = new Uint8Array(rw * rh);
    // 水平方向
    for (let y = 0; y < rh; y++) {
        const row = y * rw;
        let count = 0;
        for (let x = 0; x < Math.min(r, rw); x++) count += mask[row + x];
        for (let x = 0; x < rw; x++) {
            if (x + r < rw) count += mask[row + x + r];
            if (x - r - 1 >= 0) count -= mask[row + x - r - 1];
            if (count > 0) tmp[row + x] = 1;
        }
    }
    // 垂直方向
    const out = new Uint8Array(rw * rh);
    for (let x = 0; x < rw; x++) {
        let count = 0;
        for (let y = 0; y < Math.min(r, rh); y++) count += tmp[y * rw + x];
        for (let y = 0; y < rh; y++) {
            if (y + r < rh) count += tmp[(y + r) * rw + x];
            if (y - r - 1 >= 0) count -= tmp[(y - r - 1) * rw + x];
            if (count > 0) out[y * rw + x] = 1;
        }
    }
    return out;
}

// 連結成分ラベリング（4連結・スタック式フラッドフィル）
// 戻り値: { labels: Int32Array, comps: [{id,x0,y0,x1,y1,pix}] }
function labelComponents(mask, rw, rh) {
    const labels = new Int32Array(rw * rh);
    const stack = new Int32Array(rw * rh);
    const comps = [];
    let nextId = 1;
    for (let start = 0; start < rw * rh; start++) {
        if (mask[start] === 0 || labels[start] !== 0) continue;
        const id = nextId++;
        let sp = 0;
        stack[sp++] = start;
        labels[start] = id;
        let x0 = rw, y0 = rh, x1 = 0, y1 = 0, pix = 0;
        while (sp > 0) {
            const p = stack[--sp];
            const px = p % rw, py = (p / rw) | 0;
            pix++;
            if (px < x0) x0 = px; if (px > x1) x1 = px;
            if (py < y0) y0 = py; if (py > y1) y1 = py;
            if (px > 0 && mask[p - 1] && !labels[p - 1]) { labels[p - 1] = id; stack[sp++] = p - 1; }
            if (px < rw - 1 && mask[p + 1] && !labels[p + 1]) { labels[p + 1] = id; stack[sp++] = p + 1; }
            if (py > 0 && mask[p - rw] && !labels[p - rw]) { labels[p - rw] = id; stack[sp++] = p - rw; }
            if (py < rh - 1 && mask[p + rw] && !labels[p + rw]) { labels[p + rw] = id; stack[sp++] = p + rw; }
        }
        comps.push({ id, x0, y0, x1: x1 + 1, y1: y1 + 1, pix });
    }
    return { labels, comps };
}

// 領域を解析して連結成分に分ける
// 戻り値: { bg, labels, comps, rect } / 分解に適さない場合は null
function analyzeRegion(data, W, rect, r) {
    const rw = rect.x1 - rect.x0, rh = rect.y1 - rect.y0;
    if (rw < MIN_COMP_EDGE * 2 || rh < MIN_COMP_EDGE * 2) return null;
    const bg = dominantColor(data, W, rect);
    if (!bg || bg.ratio < MIN_BG_RATIO) return null; // 支配的な背景色が無い（写真など）
    const mask = dilate(buildMask(data, W, rect, bg), rw, rh, r);
    const { labels, comps } = labelComponents(mask, rw, rh);
    const filtered = comps.filter(c =>
        c.pix >= MIN_COMP_PIXELS &&
        (c.x1 - c.x0) >= MIN_COMP_EDGE && (c.y1 - c.y0) >= MIN_COMP_EDGE);
    if (filtered.length === 0) return null;
    return { bg, labels, comps: filtered, rect, rw, rh };
}

// =========================================================
// パーツの切り出し
// =========================================================

// 連結成分 1 個を透過 PNG 用の RGBA バッファとして切り出す
// （成分に属さない画素は透明にするので、重なっても互いを隠さない）
function cropMasked(data, W, analysis, comp) {
    const { labels, rect, rw } = analysis;
    const w = comp.x1 - comp.x0, h = comp.y1 - comp.y0;
    const out = Buffer.alloc(w * h * 4);
    for (let y = 0; y < h; y++) {
        const ly = comp.y0 + y;
        let src = ((rect.y0 + ly) * W + rect.x0 + comp.x0) * 4;
        let lab = ly * rw + comp.x0;
        let dst = y * w * 4;
        for (let x = 0; x < w; x++, src += 4, lab++, dst += 4) {
            if (labels[lab] === comp.id) {
                out[dst] = data[src]; out[dst + 1] = data[src + 1];
                out[dst + 2] = data[src + 2]; out[dst + 3] = data[src + 3] || 255;
            }
        }
    }
    return { buf: out, w, h };
}

// 矩形をそのまま切り出し、中身のパーツ部分だけ背景色で塗りつぶす
// （パネルの「背景板」を作る。sub が null なら単純な切り出し）
function cropPainted(data, W, box, sub) {
    const w = box.x1 - box.x0, h = box.y1 - box.y0;
    const out = Buffer.alloc(w * h * 4);
    for (let y = 0; y < h; y++) {
        let src = ((box.y0 + y) * W + box.x0) * 4;
        let dst = y * w * 4;
        for (let x = 0; x < w; x++, src += 4, dst += 4) {
            out[dst] = data[src]; out[dst + 1] = data[src + 1];
            out[dst + 2] = data[src + 2]; out[dst + 3] = 255;
        }
    }
    if (sub) {
        const { labels, rect, rw, bg } = sub;
        for (let ly = 0; ly < sub.rh; ly++) {
            const gy = rect.y0 + ly;
            if (gy < box.y0 || gy >= box.y1) continue;
            for (let lx = 0; lx < rw; lx++) {
                if (labels[ly * rw + lx] === 0) continue;
                const gx = rect.x0 + lx;
                if (gx < box.x0 || gx >= box.x1) continue;
                const dst = ((gy - box.y0) * w + (gx - box.x0)) * 4;
                out[dst] = bg.r; out[dst + 1] = bg.g; out[dst + 2] = bg.b; out[dst + 3] = 255;
            }
        }
    }
    return { buf: out, w, h };
}

// 全体の背景板を作る（トップレベルの全成分を背景色で塗りつぶした 1 枚）
function makeGlobalBackplate(data, W, H, analysis) {
    const out = Buffer.alloc(W * H * 4);
    for (let i = 0, o = 0; o < W * H * 4; i += 4, o += 4) {
        out[o] = data[i]; out[o + 1] = data[i + 1]; out[o + 2] = data[i + 2]; out[o + 3] = 255;
    }
    const { labels, rect, rw, bg } = analysis;
    for (let ly = 0; ly < analysis.rh; ly++) {
        for (let lx = 0; lx < rw; lx++) {
            if (labels[ly * rw + lx] === 0) continue;
            const o = ((rect.y0 + ly) * W + rect.x0 + lx) * 4;
            out[o] = bg.r; out[o + 1] = bg.g; out[o + 2] = bg.b; out[o + 3] = 255;
        }
    }
    return { buf: out, w: W, h: H };
}

// 読みやすい順（上→下、左→右）に並べるための比較関数
function readingOrder(a, b) {
    const ay = a.y0, by = b.y0;
    if (Math.abs(ay - by) > 20) return ay - by;
    return a.x0 - b.x0;
}

// =========================================================
// 1 枚の画像をパーツ一覧に分解する（分解に適さなければ null）
// 戻り値: [{x,y,w,h,buf,label}]  x,y,w,h は作業画像のピクセル座標
// =========================================================
async function segmentImage(imageBuffer) {
    let img = sharp(imageBuffer, { limitInputPixels: 268402689 });
    const meta = await img.metadata();
    if (!meta.width || !meta.height) return null;
    if (Math.min(meta.width, meta.height) < MIN_IMAGE_EDGE_PX / 4 ||
        Math.max(meta.width, meta.height) < MIN_IMAGE_EDGE_PX) return null;

    // 解析用に大きすぎる画像は縮小する（座標系はすべて作業画像基準）
    if (meta.width > MAX_WORK_EDGE_PX || meta.height > MAX_WORK_EDGE_PX) {
        img = img.resize(MAX_WORK_EDGE_PX, MAX_WORK_EDGE_PX, { fit: 'inside' });
    }
    const { data, info } = await img.ensureAlpha().raw().toBuffer({ resolveWithObject: true });
    const W = info.width, H = info.height;
    const imgArea = W * H;
    const full = { x0: 0, y0: 0, x1: W, y1: H };

    // 膨張半径を段階的に上げながら、パーツ数が上限内に収まる分解を探す
    let r = Math.max(2, Math.round(Math.min(W, H) * 0.003));
    let top = null;
    for (let attempt = 0; attempt < 3; attempt++) {
        top = analyzeRegion(data, W, full, r);
        if (!top) return null;                       // 写真など、分解に適さない
        if (top.comps.length <= MAX_PARTS) break;
        r = Math.ceil(r * 1.7);
    }
    if (!top || top.comps.length > MAX_PARTS) return null;

    // 巨大な 1 成分だけ（≒全面写真・グラデーション背景）なら分解しない
    const biggest = top.comps.reduce((m, c) => Math.max(m, (c.x1 - c.x0) * (c.y1 - c.y0)), 0);
    if (top.comps.length < 2 && biggest > imgArea * 0.92) return null;

    const parts = [];
    parts.push({ ...makeGlobalBackplate(data, W, H, top), x: 0, y: 0, label: '背景' });

    const insetPx = Math.max(6, Math.round(Math.min(W, H) * 0.012));

    // 成分列を出力する（大きなパネルは内部をさらに分解する）
    const emit = (analysis, depth) => {
        const comps = [...analysis.comps].sort(readingOrder);
        for (const c of comps) {
            if (parts.length >= MAX_PARTS + 1) return;
            // 解析領域の座標 → 画像全体の座標
            const box = {
                x0: analysis.rect.x0 + c.x0, y0: analysis.rect.y0 + c.y0,
                x1: analysis.rect.x0 + c.x1, y1: analysis.rect.y0 + c.y1,
            };
            const boxArea = (box.x1 - box.x0) * (box.y1 - box.y0);

            // パネル候補: 画像全体の 3% 以上を占め、中身がありそうな大きさ
            if (depth < MAX_DEPTH && boxArea >= imgArea * 0.03 &&
                (box.x1 - box.x0) >= 60 && (box.y1 - box.y0) >= 60 &&
                parts.length < MAX_PARTS - 2) {
                const inner = {
                    x0: box.x0 + insetPx, y0: box.y0 + insetPx,
                    x1: box.x1 - insetPx, y1: box.y1 - insetPx,
                };
                const sub = analyzeRegion(data, W, inner, Math.max(3, r >> 1));
                if (sub && sub.comps.length >= 2) {
                    const cover = sub.comps.reduce((s, sc) =>
                        s + (sc.x1 - sc.x0) * (sc.y1 - sc.y0), 0);
                    const innerArea = (inner.x1 - inner.x0) * (inner.y1 - inner.y0);
                    if (cover <= innerArea * 0.95) {
                        // パネルとして採用: 背景板 + 中身に分ける
                        parts.push({ ...cropPainted(data, W, box, sub), x: box.x0, y: box.y0, label: '枠' });
                        emit(sub, depth + 1);
                        continue;
                    }
                }
            }
            // 通常パーツ: マスク部分だけを透過 PNG として切り出す
            parts.push({ ...cropMasked(data, W, analysis, c), x: box.x0, y: box.y0, label: 'パーツ' });
        }
    };
    emit(top, 0);

    if (parts.length - 1 < MIN_PARTS) return null; // 分解する価値が無い

    // RGBA バッファ → PNG に変換
    for (const p of parts) {
        p.png = await sharp(p.buf, { raw: { width: p.w, height: p.h, channels: 4 } })
            .png().toBuffer();
        delete p.buf;
    }
    return { parts, workW: W, workH: H };
}

// =========================================================
// PPTX の走査と再構築
// =========================================================

// XML 内の数値属性をパースする小さなヘルパー
function attrInt(str, re) {
    const m = str.match(re);
    return m ? parseInt(m[1], 10) : null;
}

// スライド 1 枚を処理する。分解した場合は true を返す
async function processSlide(zip, slidePath, slideW, slideH, state) {
    const relsPath = slidePath.replace(/slides\/(slide\d+\.xml)$/, 'slides/_rels/$1.rels');
    const relsFile = zip.file(relsPath);
    const slideFile = zip.file(slidePath);
    if (!relsFile || !slideFile) return false;

    let xml = await slideFile.async('string');
    let rels = await relsFile.async('string');

    // rId → メディアパス
    const relMap = {};
    for (const m of rels.matchAll(/Id="(rId\d+)"[^>]*Target="\.\.\/media\/([^"]+)"/g)) {
        relMap[m[1]] = `ppt/media/${m[2]}`;
    }

    const pics = xml.match(/<p:pic>[\s\S]*?<\/p:pic>/g) || [];
    const slideArea = slideW * slideH;
    let changed = false;

    for (const pic of pics) {
        // 回転・反転・トリミングされた画像は対象外（座標変換が複雑になるため）
        if (/<a:xfrm[^>]*(rot|flipH|flipV)=/.test(pic) || /<a:srcRect/.test(pic)) continue;
        const offX = attrInt(pic, /<a:off x="(-?\d+)" y="-?\d+"/);
        const offY = attrInt(pic, /<a:off x="-?\d+" y="(-?\d+)"/);
        const extCx = attrInt(pic, /<a:ext cx="(\d+)" cy="\d+"/);
        const extCy = attrInt(pic, /<a:ext cx="\d+" cy="(\d+)"/);
        const rid = (pic.match(/r:embed="(rId\d+)"/) || [])[1];
        if (offX === null || offY === null || !extCx || !extCy || !rid || !relMap[rid]) continue;
        if (extCx * extCy < slideArea * PIC_AREA_RATIO_MIN) continue; // 全面画像だけが対象

        const mediaFile = zip.file(relMap[rid]);
        if (!mediaFile) continue;
        if (!/\.(png|jpe?g|gif|bmp|tiff?)$/i.test(relMap[rid])) continue;

        let seg;
        try {
            seg = await segmentImage(await mediaFile.async('nodebuffer'));
        } catch (e) {
            console.warn(`🧩 分解の解析に失敗 (${relMap[rid]}): ${e.message}`);
            continue;
        }
        if (!seg) continue; // 写真など、分解に適さない画像はそのまま

        // ピクセル座標 → EMU 座標のスケール
        const sx = extCx / seg.workW, sy = extCy / seg.workH;

        // 既存 ID の最大値を調べて新規 ID を採番する
        let maxShapeId = 0;
        for (const m of xml.matchAll(/\bid="(\d+)"/g)) maxShapeId = Math.max(maxShapeId, parseInt(m[1], 10));
        let maxRid = 0;
        for (const m of rels.matchAll(/Id="rId(\d+)"/g)) maxRid = Math.max(maxRid, parseInt(m[1], 10));

        // 各パーツをメディアとして追加し、pic 要素を生成する
        const newPics = [];
        const newRels = [];
        let partNo = 0;
        for (const p of seg.parts) {
            partNo++;
            const mediaName = `decomp${++state.mediaSeq}.png`;
            zip.file(`ppt/media/${mediaName}`, p.png);
            const newRid = `rId${++maxRid}`;
            newRels.push(`<Relationship Id="${newRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${mediaName}"/>`);
            const x = offX + Math.round(p.x * sx);
            const y = offY + Math.round(p.y * sy);
            const cx = Math.max(1, Math.round(p.w * sx));
            const cy = Math.max(1, Math.round(p.h * sy));
            const name = p.label === '背景' ? '図解の背景' : `図解${p.label} ${partNo - 1}`;
            newPics.push(
                `<p:pic><p:nvPicPr><p:cNvPr id="${++maxShapeId}" name="${name}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
                `<p:blipFill><a:blip r:embed="${newRid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
                `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
                `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
            );
        }

        // 元の 1 枚画像を、パーツ群（Z順: 背景板 → 各パーツ）で置き換える
        xml = xml.replace(pic, newPics.join(''));
        rels = rels.replace('</Relationships>', newRels.join('') + '</Relationships>');
        changed = true;
        console.log(`🧩 一枚絵を検出: ${slidePath} → ${seg.parts.length - 1} パーツに分解`);
    }

    if (changed) {
        zip.file(slidePath, xml);
        zip.file(relsPath, rels);
    }
    return changed;
}

// =========================================================
// エントリポイント: PPTX 内の全スライドを対象に分解を試みる
// 既存機能を壊さないため、どんな失敗でも例外を外に漏らさない
// =========================================================
export async function decomposeFlatImages(pptxPath) {
    try {
        const zip = await JSZip.loadAsync(fs.readFileSync(pptxPath));

        // スライドサイズ
        const presFile = zip.file('ppt/presentation.xml');
        if (!presFile) return;
        const pres = await presFile.async('string');
        const sldSz = pres.match(/<p:sldSz cx="(\d+)" cy="(\d+)"/);
        if (!sldSz) return;
        const slideW = parseInt(sldSz[1], 10), slideH = parseInt(sldSz[2], 10);

        const slidePaths = Object.keys(zip.files)
            .filter(p => /^ppt\/slides\/slide\d+\.xml$/.test(p))
            .sort((a, b) => parseInt(a.match(/\d+/g).pop(), 10) - parseInt(b.match(/\d+/g).pop(), 10));

        const state = { mediaSeq: 0 };
        let anyChanged = false;
        for (const slidePath of slidePaths) {
            if (await processSlide(zip, slidePath, slideW, slideH, state)) anyChanged = true;
        }
        if (!anyChanged) return;

        // PNG の Content-Type 定義が無ければ追加する
        const ctFile = zip.file('[Content_Types].xml');
        if (ctFile) {
            let ct = await ctFile.async('string');
            if (!/Extension="png"/.test(ct)) {
                ct = ct.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>');
                zip.file('[Content_Types].xml', ct);
            }
        }

        const content = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
        fs.writeFileSync(pptxPath, content);
    } catch (e) {
        // 分解はあくまで付加機能。失敗しても従来の変換結果をそのまま返す
        console.warn(`🧩 一枚絵分解をスキップ: ${e.message}`);
    }
}
