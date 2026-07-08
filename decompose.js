// =========================================================
// 一枚絵（フラット化された図解）のパーツ分解 + OCR テキスト化
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
//   3. 大きなかたまり（パネルやカード）は内部をさらに再帰的に
//      分解し、「背景板」と「中身の物体」に分ける。写真のような
//      グラデーション領域は再分解せず 1 つの物体として保つ。
//   4. 文字らしいパーツは OCR (tesseract) にかけ、確信度が
//      高い場合だけ編集可能なテキストボックスに置き換える。
//      確信度が低ければ画像のまま残す（誤読テキストを防ぐ）。
//   5. スライドを「きれいな背景板 + 各パーツ」で再構築する。
//
// 完全に元の描画命令へ戻す正攻法は存在しないため、これは
// 「多少の劣化を許容して編集可能性を得る」近似的なアプローチ。
// 判定に失敗した場合は必ず元のスライドを無傷で残す。
// =========================================================
import { createRequire } from 'module';
const require = createRequire(import.meta.url);

const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { execFile } = require('child_process');
const sharp = require('sharp');
const JSZip = require('jszip');

// ---- 分解のチューニングパラメータ ----
const PIC_AREA_RATIO_MIN = 0.60;   // スライド面積に対する画像の占有率（これ以上で分解候補）
const MIN_IMAGE_EDGE_PX = 400;     // 分解対象とする画像の最小辺長
const MAX_WORK_EDGE_PX = 3200;     // 解析時の最大辺長（超える場合は縮小して処理）
const COLOR_DIST_T = 70;           // 背景色との距離しきい値 (|Δr|+|Δg|+|Δb|)
const MIN_BG_RATIO = 0.25;         // 背景色の占有率がこれ未満なら「写真」とみなし分解しない
const MIN_COMP_PIXELS = 30;        // これより小さい連結成分はノイズとして背景板に残す
const MIN_COMP_EDGE = 5;           // 連結成分バウンディングボックスの最小辺 (px)
const MAX_PARTS = 240;             // 1画像あたりの最大パーツ数
const MIN_PARTS = 4;               // これ未満しか分解できない場合は元のまま残す
const MAX_DEPTH = 4;               // パネル内再分解の最大深さ
const CONTAINER_AREA_RATIO = 0.012;// 内部を再分解する候補の最小面積（画像全体比）
const CONTAINER_MIN_W = 80;        // 内部を再分解する候補の最小幅 (px)
const CONTAINER_MIN_H = 110;       // 同・最小高さ (px)（文字行・帯を誤って割らないため大きめ）
const TEXT_MERGE_MAX_H = 90;       // 文字行マージの対象になる成分の最大高さ (px)

// ---- OCR のチューニングパラメータ ----
const OCR_MIN_WORD_CONF = 70;      // 単語の最低確信度（未満の内容語が混ざる行は不採用）
const OCR_MIN_LINE_CONF = 82;      // 行の平均確信度がこれ未満なら不採用
const OCR_MIN_COVERAGE = 0.45;     // 採用行の面積がパーツ面積のこれ未満なら不採用（読み漏らし防止）
const OCR_MIN_PART_H = 12;         // OCR にかけるパーツの最小高さ (px)（小さすぎると誤読が増える）
const OCR_MAX_PARTS = 120;         // 1画像あたりの OCR 試行数上限
const OCR_TOTAL_BUDGET_MS = 120000;// 1画像あたりの OCR 合計時間の上限
const OCR_CONCURRENCY = 4;         // tesseract の並列実行数

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

// 領域を解析して連結成分に分ける（分解に適さない場合は null）
// bgOverride を渡すと、領域内の最頻色ではなくその色を背景として使う
// （写真が大きく背景色が最頻にならないパネルの再分解用）
function analyzeRegion(data, W, rect, r, bgOverride) {
    const rw = rect.x1 - rect.x0, rh = rect.y1 - rect.y0;
    if (rw < MIN_COMP_EDGE * 2 || rh < MIN_COMP_EDGE * 2) return null;
    const bg = bgOverride || dominantColor(data, W, rect);
    if (!bg) return null;
    if (!bgOverride && bg.ratio < MIN_BG_RATIO) return null; // 支配的な背景色が無い（写真など）
    const mask = dilate(buildMask(data, W, rect, bg), rw, rh, r);
    const { labels, comps } = labelComponents(mask, rw, rh);
    const filtered = comps.filter(c =>
        c.pix >= MIN_COMP_PIXELS &&
        (c.x1 - c.x0) >= MIN_COMP_EDGE && (c.y1 - c.y0) >= MIN_COMP_EDGE);
    if (filtered.length === 0) return null;
    return { bg, labels, comps: filtered, rect, rw, rh };
}

// 同じ行に並ぶ文字らしい成分をひとつの「文字行」にまとめる
// （文字単位のバラバラな成分を、動かしやすく OCR しやすい行単位へ）
function mergeTextRuns(comps) {
    const items = comps.map(c => ({ ...c, ids: c.ids || [c.id] }));
    items.sort((a, b) => a.x0 - b.x0);
    let mergedAny = true;
    while (mergedAny) {
        mergedAny = false;
        for (let i = 0; i < items.length; i++) {
            const a = items[i];
            const ah = a.y1 - a.y0;
            if (ah < 8 || ah > TEXT_MERGE_MAX_H) continue;
            for (let j = i + 1; j < items.length; j++) {
                const b = items[j];
                const bh = b.y1 - b.y0;
                if (bh < 8 || bh > TEXT_MERGE_MAX_H) continue;
                // 高さが大きく違うもの（アイコンと文字など）はマージしない
                if (Math.min(ah, bh) / Math.max(ah, bh) < 0.55) continue;
                // 垂直方向の重なりが小さい・横の間隔が広いものはマージしない
                const overlap = Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0);
                if (overlap < 0.7 * Math.min(ah, bh)) continue;
                const gap = Math.max(a.x0, b.x0) - Math.min(a.x1, b.x1);
                if (gap > 0.5 * Math.min(ah, bh)) continue;
                a.x0 = Math.min(a.x0, b.x0); a.y0 = Math.min(a.y0, b.y0);
                a.x1 = Math.max(a.x1, b.x1); a.y1 = Math.max(a.y1, b.y1);
                a.pix += b.pix;
                a.ids.push(...b.ids);
                items.splice(j, 1);
                mergedAny = true;
                j--;
            }
        }
    }
    return items;
}

// =========================================================
// パーツの切り出し（RGBA バッファ + 統計情報）
// =========================================================

// 連結成分 1 個を透過付き RGBA バッファとして切り出す
// 成分に属さない画素は透明にする。同時に OCR 判定用の統計も計算する。
// マスクは膨張済みなので背景色のハロー（文字周りの縁）を含む。
// 統計はハローに騙されないよう、背景色から離れた「コア画素」で取る
// claimed: 出力済み画素のマスク（画像全体）。同じ画素が別の階層から
// 二重にパーツ化されるのを防ぐ。切り出した画素は claimed に記録する
function cropMasked(data, W, analysis, comp, claimed) {
    const { labels, rect, rw, bg } = analysis;
    const idSet = new Set(comp.ids || [comp.id]); // 文字行マージで複数 ID を持つことがある
    const w = comp.x1 - comp.x0, h = comp.y1 - comp.y0;
    const out = Buffer.alloc(w * h * 4);
    const hist = new Uint32Array(4096);
    let fgPix = 0, corePix = 0, skipped = 0;
    for (let y = 0; y < h; y++) {
        const ly = comp.y0 + y;
        let src = ((rect.y0 + ly) * W + rect.x0 + comp.x0) * 4;
        let lab = ly * rw + comp.x0;
        let dst = y * w * 4;
        for (let x = 0; x < w; x++, src += 4, lab++, dst += 4) {
            if (idSet.has(labels[lab])) {
                const g = src >> 2; // 画像全体での画素番号
                if (claimed[g]) { skipped++; continue; } // 既に別パーツが持っている
                claimed[g] = 1;
                out[dst] = data[src]; out[dst + 1] = data[src + 1];
                out[dst + 2] = data[src + 2]; out[dst + 3] = data[src + 3] || 255;
                fgPix++;
                const d = Math.abs(out[dst] - bg.r) + Math.abs(out[dst + 1] - bg.g) + Math.abs(out[dst + 2] - bg.b);
                if (d > COLOR_DIST_T) {
                    hist[((out[dst] >> 4) << 8) | ((out[dst + 1] >> 4) << 4) | (out[dst + 2] >> 4)]++;
                    corePix++;
                }
            }
        }
    }
    // ほとんどの画素が出力済みなら、このパーツは重複なので出さない
    if (fgPix < (fgPix + skipped) * 0.15 || fgPix < MIN_COMP_PIXELS) return null;
    // コア画素の上位 2 色とその占有率
    let k1 = -1, n1 = 0, k2 = -1, n2 = 0;
    for (let k = 0; k < 4096; k++) {
        if (hist[k] > n1) { k2 = k1; n2 = n1; k1 = k; n1 = hist[k]; }
        else if (hist[k] > n2) { k2 = k; n2 = hist[k]; }
    }
    const binColor = (k) => ({
        r: ((k >> 8) & 15) * 16 + 8, g: ((k >> 4) & 15) * 16 + 8, b: (k & 15) * 16 + 8,
    });
    // 白抜き文字の検出。文字は 2 通りの形で現れる:
    //   a. 「内側の穴」= 前景に挟まれた透明画素（太い文字）
    //   b. 塗り色から大きく離れた色の前景画素（細い文字は膨張処理で
    //      マスクに飲み込まれ、白い画素として塗りの中に残る）
    const t1 = k1 >= 0 ? binColor(k1) : null;
    let holePix = 0, hr = 0, hg = 0, hb = 0;
    for (let y = 0; y < h; y++) {
        const row = y * w * 4;
        let first = -1, last = -1;
        for (let x = 0; x < w; x++) if (out[row + x * 4 + 3] > 0) { if (first < 0) first = x; last = x; }
        if (first < 0) continue;
        for (let x = first + 1; x < last; x++) {
            const o = row + x * 4;
            if (out[o + 3] === 0) {
                holePix++;
                const src = ((rect.y0 + comp.y0 + y) * W + rect.x0 + comp.x0 + x) * 4;
                hr += data[src]; hg += data[src + 1]; hb += data[src + 2];
            } else if (t1) {
                const d = Math.abs(out[o] - t1.r) + Math.abs(out[o + 1] - t1.g) + Math.abs(out[o + 2] - t1.b);
                if (d > 150) {
                    holePix++;
                    hr += out[o]; hg += out[o + 1]; hb += out[o + 2];
                }
            }
        }
    }
    return {
        buf: out, w, h,
        levelBg: bg, // このパーツが乗っていた背景色（照合時のハロー除去に使う）
        stats: {
            fgPix,
            fillRatio: fgPix / (w * h),                 // マスク全体（ハロー込み）の占有率
            coreFillRatio: corePix / (w * h),           // 実際に色の付いた画素の占有率
            top1: k1 >= 0 ? { ...binColor(k1), ratio: n1 / Math.max(1, corePix) } : null,
            top2ratio: (n1 + n2) / Math.max(1, corePix),
            holeRatio: holePix / (w * h),
            holeColor: holePix > 0
                ? { r: Math.round(hr / holePix), g: Math.round(hg / holePix), b: Math.round(hb / holePix) }
                : null,
        },
    };
}

// 矩形をそのまま切り出し、「実際に出力される中身」だけを背景色で塗りつぶす
// （パネルの「背景板」を作る。paintIds に無い成分は塗らずに残す —
//   塗り消したのに出力されない成分があると表示が欠けるため）
function cropPainted(data, W, box, sub, paintIds) {
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
                const id = labels[ly * rw + lx];
                if (id === 0 || (paintIds && !paintIds.has(id))) continue;
                const gx = rect.x0 + lx;
                if (gx < box.x0 || gx >= box.x1) continue;
                const dst = ((gy - box.y0) * w + (gx - box.x0)) * 4;
                out[dst] = bg.r; out[dst + 1] = bg.g; out[dst + 2] = bg.b; out[dst + 3] = 255;
            }
        }
    }
    return { buf: out, w, h };
}

// 全体の背景板を作る（実際に出力されるトップレベル成分を塗り消した 1 枚）
function makeGlobalBackplate(data, W, H, analysis, paintIds) {
    const out = Buffer.alloc(W * H * 4);
    for (let i = 0, o = 0; o < W * H * 4; i += 4, o += 4) {
        out[o] = data[i]; out[o + 1] = data[i + 1]; out[o + 2] = data[i + 2]; out[o + 3] = 255;
    }
    const { labels, rect, rw, bg } = analysis;
    for (let ly = 0; ly < analysis.rh; ly++) {
        for (let lx = 0; lx < rw; lx++) {
            const id = labels[ly * rw + lx];
            if (id === 0 || (paintIds && !paintIds.has(id))) continue;
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
// =========================================================
export async function segmentImage(imageBuffer) {
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

    const insetPx = Math.max(6, Math.round(Math.min(W, H) * 0.012));
    const claimed = new Uint8Array(W * H); // 二重出力防止用の出力済み画素マスク

    // =====================================================
    // フェーズ1: 出力計画を立てる
    //
    // 大原則:「背景板から塗り消すのは、実際にパーツとして出力する
    // 成分だけ」。塗り消したのに出力しない成分があると、その部分の
    // 表示が欠けてしまう。そのため先に出力する成分を確定してから
    // （フェーズ2で）その成分だけを塗り消して描画する。
    //   - ノイズとして除外した成分 → 塗り消さない（背景板に残る）
    //   - 容器の境界にかかる成分   → 出力せず塗り消さない
    //     （容器の内側しか解析できず、半分だけのパーツになるため）
    // =====================================================
    let budget = MAX_PARTS;
    const makePlan = (analysis, depth, isInner) => {
        let comps = mergeTextRuns(analysis.comps).sort(readingOrder);
        // 容器の inner 境界に接している成分は境界の外に続きがある
        // 可能性が高いので、出力せず容器の背景板に残す
        if (isInner) {
            comps = comps.filter(c =>
                c.x0 > 1 && c.y0 > 1 && c.x1 < analysis.rw - 1 && c.y1 < analysis.rh - 1);
        }

        // 容器（枠+中身がひとかたまりの成分）の判定
        const containers = new Map(); // comp → { box, inner, sub }
        for (const c of comps) {
            const box = {
                x0: analysis.rect.x0 + c.x0, y0: analysis.rect.y0 + c.y0,
                x1: analysis.rect.x0 + c.x1, y1: analysis.rect.y0 + c.y1,
            };
            const boxArea = (box.x1 - box.x0) * (box.y1 - box.y0);
            if (depth >= MAX_DEPTH || boxArea < imgArea * CONTAINER_AREA_RATIO ||
                (box.x1 - box.x0) < CONTAINER_MIN_W || (box.y1 - box.y0) < CONTAINER_MIN_H ||
                budget < 3) continue;
            const inner = {
                x0: box.x0 + insetPx, y0: box.y0 + insetPx,
                x1: box.x1 - insetPx, y1: box.y1 - insetPx,
            };
            const innerArea = (inner.x1 - inner.x0) * (inner.y1 - inner.y0);
            // 数パターンの解析を試し、うまく中身が分かれるものを採用する
            //   1. 通常（領域内の最頻色を背景に）
            //   2. 膨張半径を最小にして密集した中身を切り離す
            //   3. 親レベルの背景色を強制（写真が大きく最頻色が背景でないパネル）
            const tries = [
                { r: Math.max(3, r >> 1) },
                { r: 2 },
                { r: Math.max(3, r >> 1), bg: analysis.bg },
                { r: 2, bg: analysis.bg },
            ];
            let sub = null;
            for (const t of tries) {
                const s = analyzeRegion(data, W, inner, t.r, t.bg);
                if (!s || s.comps.length < 2) continue;
                // 中身の画素がぎっしり詰まっている領域（写真・グラデーション）は分解しない
                const inkPix = s.comps.reduce((sum, sc) => sum + sc.pix, 0);
                if (inkPix > innerArea * 0.90) continue;
                // ほぼ全面を占める成分が残る分解は「分けられていない」ので不採用
                // （採用すると同じ枠を一回り小さく切り直すだけの入れ子が生まれる）
                if (s.comps.some(sc =>
                    (sc.x1 - sc.x0) * (sc.y1 - sc.y0) > innerArea * 0.90)) continue;
                sub = s;
                break;
            }
            if (!sub) {
                if (process.env.DECOMP_DEBUG) console.log(`  容器却下 ${box.x0},${box.y0} ${box.x1-box.x0}x${box.y1-box.y0}`);
                continue;
            }
            if (process.env.DECOMP_DEBUG) console.log(`  容器採用 ${box.x0},${box.y0} ${box.x1-box.x0}x${box.y1-box.y0}: sub=${sub.comps.length}`);
            containers.set(c, { box, inner, sub });
        }

        // ある成分の画素の過半が「別の」容器の内側にあるか
        // （その成分は容器の再分解側から出力されるので、ここでは出力しない）
        const coveredByOtherContainer = (self, box) => {
            const idSet = new Set(self.ids || [self.id]);
            for (const [cc, k] of containers) {
                if (cc === self) continue;
                if (box.x1 <= k.inner.x0 || box.x0 >= k.inner.x1 ||
                    box.y1 <= k.inner.y0 || box.y0 >= k.inner.y1) continue;
                let inside = 0;
                for (let y = self.y0; y < self.y1; y++) {
                    const gy = analysis.rect.y0 + y;
                    const row = y * analysis.rw;
                    for (let x = self.x0; x < self.x1; x++) {
                        if (!idSet.has(analysis.labels[row + x])) continue;
                        const gx = analysis.rect.x0 + x;
                        if (gx >= k.inner.x0 && gx < k.inner.x1 &&
                            gy >= k.inner.y0 && gy < k.inner.y1) inside++;
                    }
                }
                if (inside >= self.pix * 0.5) return true;
            }
            return false;
        };

        // 出力計画の作成
        const node = { analysis, entries: [], paintIds: new Set() };
        for (const c of comps) {
            if (budget <= 0) break;
            const box = {
                x0: analysis.rect.x0 + c.x0, y0: analysis.rect.y0 + c.y0,
                x1: analysis.rect.x0 + c.x1, y1: analysis.rect.y0 + c.y1,
            };
            if (coveredByOtherContainer(c, box)) {
                // 覆っている容器の背景板の下に隠れるので塗り消してよい
                for (const id of (c.ids || [c.id])) node.paintIds.add(id);
                continue;
            }
            const k = containers.get(c);
            if (k) {
                budget--;
                const subNode = makePlan(k.sub, depth + 1, true);
                if (subNode.entries.length > 0) {
                    node.entries.push({ type: 'panel', comp: c, box: k.box, sub: k.sub, subNode });
                } else {
                    budget++; // 中身が出せないなら普通のパーツとして出す
                    node.entries.push({ type: 'piece', comp: c, box });
                    budget--;
                }
            } else {
                budget--;
                node.entries.push({ type: 'piece', comp: c, box });
            }
            for (const id of (c.ids || [c.id])) node.paintIds.add(id);
        }
        return node;
    };
    const rootNode = makePlan(top, 0, false);

    // =====================================================
    // フェーズ2: 計画どおりに描画する
    // =====================================================
    const parts = [];
    parts.push({ ...makeGlobalBackplate(data, W, H, top, rootNode.paintIds), x: 0, y: 0, kind: 'backplate' });
    const render = (node) => {
        for (const e of node.entries) {
            if (e.type === 'panel') {
                parts.push({
                    ...cropPainted(data, W, e.box, e.sub, e.subNode.paintIds),
                    x: e.box.x0, y: e.box.y0, kind: 'panel',
                });
                render(e.subNode);
            } else {
                const crop = cropMasked(data, W, node.analysis, e.comp, claimed);
                if (crop) parts.push({ ...crop, x: e.box.x0, y: e.box.y0, kind: 'piece' });
            }
        }
    };
    render(rootNode);

    if (parts.length - 1 < MIN_PARTS) return null; // 分解する価値が無い

    return { parts, workW: W, workH: H };
}

// =========================================================
// OCR（tesseract CLI）
// =========================================================

// tesseract が使えるかどうか（初回だけ実行して結果をキャッシュ）
let tesseractProbe = null;
function isTesseractAvailable() {
    if (!tesseractProbe) {
        tesseractProbe = new Promise((resolve) => {
            execFile('tesseract', ['--version'], { timeout: 10000 }, (err) => resolve(!err));
        });
    }
    return tesseractProbe;
}

// OCR 対象らしいパーツかどうかの判定
//   normal  : 透明背景に文字が描かれたパーツ（文字色 = 前景色）
//   inverse : 塗りつぶし図形に白抜き文字（文字 = 内側の透明な穴）
export function classifyForOcr(part, imgW, imgH) {
    if (part.kind !== 'piece' || !part.stats) return null;
    const { w, h, stats } = part;
    if (h < OCR_MIN_PART_H || w < 14) return null;
    if (h > imgH * 0.3 || w * h > imgW * imgH * 0.2) return null;
    if (stats.coreFillRatio < 0.008) return null;
    // 白抜き文字: 色の付いた画素が支配的で、内側に穴（文字の抜き）がある
    if (stats.coreFillRatio >= 0.55 && stats.holeRatio >= 0.02 && stats.holeColor) {
        return 'inverse';
    }
    // 通常文字: 背景の上に文字・記号が描かれている
    // （アイコン等も候補になるが、OCR の確信度ゲートで自然に除外される）
    if (stats.coreFillRatio <= 0.60) {
        return 'normal';
    }
    return null;
}

// パーツから OCR 用のモノクロ画像を作る（拡大 + 白地に黒文字へ正規化）
async function buildOcrImage(part, mode) {
    const scale = part.h < 24 ? 4 : 3;
    const raw = buildOcrRaw(part, mode);
    const png = await sharp(raw, { raw: { width: part.w, height: part.h, channels: 4 } })
        .resize({ width: part.w * scale, kernel: 'lanczos3' })
        .grayscale().normalise()
        .png().toBuffer();
    return { png, scale, w: part.w * scale, h: Math.round(part.h * scale) };
}

// OCR 用の RGBA バッファを作る（白地に黒文字へ正規化、拡大前）
function buildOcrRaw(part, mode) {
    const raw = Buffer.alloc(part.w * part.h * 4);
    if (mode === 'inverse') {
        // 塗り部分は元の色、内側の穴（文字）は白にしてから全体を反転する。
        // 塗りと文字の境界のアンチエイリアスが保たれ、OCR 精度が上がる
        const bar = part.stats.top1 || { r: 0, g: 0, b: 0 };
        for (let y = 0; y < part.h; y++) {
            const row = y * part.w * 4;
            let first = -1, last = -1;
            for (let x = 0; x < part.w; x++) if (part.buf[row + x * 4 + 3] > 0) { if (first < 0) first = x; last = x; }
            for (let x = 0; x < part.w; x++) {
                const o = row + x * 4;
                let r, g, b;
                if (part.buf[o + 3] > 0) {
                    r = part.buf[o]; g = part.buf[o + 1]; b = part.buf[o + 2];
                } else if (first >= 0 && x > first && x < last) {
                    r = 255; g = 255; b = 255;      // 内側の穴 = 文字
                } else {
                    r = bar.r; g = bar.g; b = bar.b; // 外側は塗り色と同化させる
                }
                raw[o] = 255 - r; raw[o + 1] = 255 - g; raw[o + 2] = 255 - b; raw[o + 3] = 255;
            }
        }
    } else {
        // 透明 → 白地、前景はそのまま（色文字はグレースケール化される）
        for (let y = 0; y < part.h; y++) {
            for (let x = 0; x < part.w; x++) {
                const o = (y * part.w + x) * 4;
                if (part.buf[o + 3] === 0) {
                    raw[o] = 255; raw[o + 1] = 255; raw[o + 2] = 255; raw[o + 3] = 255;
                } else {
                    raw[o] = part.buf[o]; raw[o + 1] = part.buf[o + 1];
                    raw[o + 2] = part.buf[o + 2]; raw[o + 3] = 255;
                }
            }
        }
    }
    return raw;
}

// tesseract を実行して TSV を返す
function runTesseract(pngPath) {
    return new Promise((resolve) => {
        execFile('tesseract', [pngPath, 'stdout', '-l', 'jpn+eng', '--psm', '6', 'tsv'],
            { timeout: OCR_TOTAL_BUDGET_MS, maxBuffer: 16 * 1024 * 1024 },
            (err, stdout) => resolve(err ? null : stdout));
    });
}

// 記号・罫線くずだけのトークン（枠線の切れ端などの誤検出）
// ※ +−% などの数値に付く記号は内容として保持する
function isGarbageToken(text) {
    return !/[\p{L}\p{N}⺀-鿿぀-ヿ０-ｚ%％+＋±\-−]/u.test(text);
}

// 全角文字の間に OCR が入れた余計なスペースを取り除く
function tidyJapanese(text) {
    return text
        .replace(/([⺀-鿿　-ヿ＀-￯]) +(?=[⺀-鿿　-ヿ＀-￯])/g, '$1')
        .trim();
}

// 行の集合に確信度ゲートをかけ、採用できる行だけを返す
// 「内容語があるのに確信度が低い行」が 1 つでもあれば null
// （誤読テキストを出すくらいなら、パーツ全体を画像のまま残す）
function gateLines(lineList, partW, partH, isDiscLike, maxLineH) {
    const accepted = [];
    for (const words of lineList) {
        let content = words.filter(w => !isGarbageToken(w.text));
        if (content.length === 0) continue; // 記号くずだけの行は無視

        // 小さすぎる文字は誤読が多いので信頼しない（画像のまま残る）。
        // 大きな見出し文字も誤字が目立ちすぎるため画像のまま残す
        // （タイトルは動かせれば十分で、打ち替えの需要は本文より低い）
        const lineH = Math.max(...content.map(w => w.y1 - w.y0));
        if (lineH < 12 || (maxLineH && lineH > maxLineH)) continue;

        // 行頭のバッジ（❶❷…の丸数字など）は高確信度で誤読されやすい。
        // 「最初の大きな空白より前にある短いトークン群」が円盤状
        // （四隅が空で中心が濃い）なら行から外す（バッジ自体は画像の
        // まま残り、本文だけがテキスト化される）
        if (content.length >= 2 && isDiscLike) {
            let k = 0; // 最初の大きな空白の位置
            while (k + 1 < content.length &&
                   content[k + 1].x0 - content[k].x1 < lineH * 0.3) k++;
            const lead = content.slice(0, k + 1);
            const rest = content.slice(k + 1);
            const restText = tidyJapanese(rest.map(w => w.text).join(' '));
            const leadText = lead.map(w => w.text).join('');
            const leadBox = {
                x0: Math.min(...lead.map(w => w.x0)), y0: Math.min(...lead.map(w => w.y0)),
                x1: Math.max(...lead.map(w => w.x1)), y1: Math.max(...lead.map(w => w.y1)),
            };
            const doStrip = rest.length > 0 && leadText.length <= 2 &&
                (leadBox.x1 - leadBox.x0) <= lineH * 1.9 &&
                restText.length >= 3 &&
                isDiscLike(leadBox);
            if (process.env.DECOMP_DEBUG && rest.length > 0 && leadText.length <= 3) {
                console.log(`  バッジ判定 lead="${leadText}" rest="${restText.slice(0, 10)}" ` +
                    `w=${(leadBox.x1 - leadBox.x0).toFixed(0)} lineH=${lineH.toFixed(0)} ` +
                    `disc=${isDiscLike(leadBox)} strip=${doStrip}`);
            }
            if (doStrip) content = rest;
        }
        if (content.some(w => w.conf < OCR_MIN_WORD_CONF)) return null;
        const mean = content.reduce((s, w) => s + w.conf, 0) / content.length;
        if (mean < OCR_MIN_LINE_CONF) return null;
        const text = tidyJapanese(content.map(w => w.text).join(' '));
        if (!text) continue;
        accepted.push({
            text,
            x0: Math.min(...content.map(w => w.x0)), y0: Math.min(...content.map(w => w.y0)),
            x1: Math.max(...content.map(w => w.x1)), y1: Math.max(...content.map(w => w.y1)),
        });
    }
    if (accepted.length === 0) return null;
    // 読み漏らし防止: 採用行の合計面積がパーツ面積に対して小さすぎる場合、
    // 認識できていない文字・図形が残っているとみなし、パーツ全体を画像のまま残す
    const covered = accepted.reduce((s, l) => s + (l.x1 - l.x0) * (l.y1 - l.y0), 0);
    if (covered < partW * partH * OCR_MIN_COVERAGE) return null;
    return accepted;
}

// =========================================================
// 認識テキストの再描画照合
//
// OCR は「高い確信度で誤読する」ことがある（例: 2→ノ, ②→の）。
// 認識したテキストをフォントで実際に描画し、元の字形と
// 「縦横比」「インクの列プロファイル」を比較して、形が合わない
// 認識結果を捨てる。字形レベルの検証なので、確信度だけでは
// 拾えない誤読・読み漏らしをブロックできる
// =========================================================

// テキストを描画してインク（黒画素）の列プロファイルと bbox を得る
async function renderTextInk(text, bold) {
    const fsPx = 48;
    const w = Math.max(64, Math.ceil(fsPx * 1.1 * (text.length + 1)));
    const h = Math.ceil(fsPx * 1.8);
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}">` +
        `<rect width="100%" height="100%" fill="white"/>` +
        `<text x="8" y="${Math.round(fsPx * 1.25)}" font-family="Noto Sans CJK JP, Noto Sans JP, sans-serif"` +
        ` font-size="${fsPx}" font-weight="${bold ? 700 : 400}" fill="black">${escapeXml(text)}</text></svg>`;
    const { data, info } = await sharp(Buffer.from(svg))
        .flatten({ background: '#ffffff' }).grayscale().raw()
        .toBuffer({ resolveWithObject: true });
    return inkProfile((x, y) => data[y * info.width + x] < 128, info.width, info.height);
}

// インク判定関数から bbox と列プロファイル（48分割の正規化インク量）を作る
function inkProfile(isInk, w, h) {
    let x0 = w, y0 = h, x1 = -1, y1 = -1;
    for (let y = 0; y < h; y++) {
        for (let x = 0; x < w; x++) {
            if (!isInk(x, y)) continue;
            if (x < x0) x0 = x; if (x > x1) x1 = x;
            if (y < y0) y0 = y; if (y > y1) y1 = y;
        }
    }
    if (x1 < 0) return null;
    const bins = new Float64Array(48);
    const bw = (x1 - x0 + 1) / 48;
    for (let y = y0; y <= y1; y++) {
        for (let x = x0; x <= x1; x++) {
            if (isInk(x, y)) bins[Math.min(47, Math.floor((x - x0) / bw))]++;
        }
    }
    let max = 0;
    for (const v of bins) max = Math.max(max, v);
    if (max > 0) for (let i = 0; i < 48; i++) bins[i] /= max;
    return { x0, y0, x1, y1, w: x1 - x0 + 1, h: y1 - y0 + 1, bins };
}

// 2 つの列プロファイルの相関係数
function profileCorrelation(a, b) {
    let ma = 0, mb = 0;
    for (let i = 0; i < 48; i++) { ma += a[i]; mb += b[i]; }
    ma /= 48; mb /= 48;
    let cov = 0, va = 0, vb = 0;
    for (let i = 0; i < 48; i++) {
        cov += (a[i] - ma) * (b[i] - mb);
        va += (a[i] - ma) ** 2; vb += (b[i] - mb) ** 2;
    }
    if (va === 0 || vb === 0) return 0;
    return cov / Math.sqrt(va * vb);
}

// CJK フォントが描画できる環境かどうか（初回だけ確認）
let cjkRenderProbe = null;
function canRenderCjk() {
    if (!cjkRenderProbe) {
        cjkRenderProbe = renderTextInk('永続', false)
            .then(p => !!p && p.w > 20).catch(() => false);
    }
    return cjkRenderProbe;
}

// パーツ内の行領域のインクプロファイルを作る
function lineInkProfile(part, mode, line) {
    const x0 = Math.max(0, Math.floor(line.x0) - 1), x1 = Math.min(part.w, Math.ceil(line.x1) + 1);
    const y0 = Math.max(0, Math.floor(line.y0) - 1), y1 = Math.min(part.h, Math.ceil(line.y1) + 1);
    // 行ごとの「内側の穴」判定のため、行範囲で前景の左右端を求めておく
    // （前景がまばらな行は穴判定の対象外にする）
    const rowSpan = [];
    for (let y = y0; y < y1; y++) {
        let first = -1, last = -1, fgCount = 0;
        const row = y * part.w * 4;
        for (let x = 0; x < part.w; x++) {
            if (part.buf[row + x * 4 + 3] > 0) { if (first < 0) first = x; last = x; fgCount++; }
        }
        if (first >= 0 && fgCount < (last - first + 1) * 0.5) { first = -1; last = -1; }
        rowSpan[y - y0] = [first, last];
    }
    const bg = part.levelBg || { r: 255, g: 255, b: 255 };
    const bar = part.stats && part.stats.top1;
    const isInk = (x, y) => {
        const gx = x + x0, gy = y + y0;
        const o = (gy * part.w + gx) * 4;
        if (mode === 'inverse') {
            const [f, l] = rowSpan[gy - y0] || [-1, -1];
            if (f < 0 || gx <= f || gx >= l) return false;
            if (part.buf[o + 3] === 0) return true; // 穴 = 文字
            if (!bar) return false;
            // 塗り色から大きく離れた画素（飲み込まれた文字）もインク
            const d = Math.abs(part.buf[o] - bar.r) + Math.abs(part.buf[o + 1] - bar.g) + Math.abs(part.buf[o + 2] - bar.b);
            return d > 150;
        }
        if (part.buf[o + 3] === 0) return false;
        // ハロー（背景色の縁）はインクに数えない
        const d = Math.abs(part.buf[o] - bg.r) + Math.abs(part.buf[o + 1] - bg.g) + Math.abs(part.buf[o + 2] - bg.b);
        return d > COLOR_DIST_T * 0.7;
    };
    return inkProfile(isInk, x1 - x0, y1 - y0);
}

// 行の認識結果を再描画照合する（true = 合格）
async function verifyLine(part, mode, line, bold) {
    const orig = lineInkProfile(part, mode, line);
    if (!orig || orig.h < 5) return null;
    // CJK フォントが無い環境では検証できないため、確信度ゲートのみに委ねる
    if (/[⺀-鿿぀-ヿ]/.test(line.text) && !(await canRenderCjk())) return {};
    const rendered = await renderTextInk(line.text, bold);
    if (!rendered) return null;
    // 縦横比の比較（フォント差を考慮して緩めに）
    const arOrig = orig.w / orig.h, arRend = rendered.w / rendered.h;
    const ratio = arRend / arOrig;
    const corr = profileCorrelation(orig.bins, rendered.bins);
    if (process.env.DECOMP_DEBUG) {
        console.log(`  照合 "${line.text}" ratio=${ratio.toFixed(2)} corr=${corr.toFixed(2)} ` +
            `orig=${orig.w}x${orig.h} rend=${rendered.w}x${rendered.h}`);
    }
    // 日本語の図解では長体（横に潰した文字）が多用されるため、
    // 「描画した方が横に長い」方向へは大きく許容する
    if (ratio < 0.60 || ratio > 2.30) return null;
    // 列プロファイルの比較
    if (corr < 0.45) return null;
    // 合格。フォントサイズ推定（幅合わせ）のため描画時の寸法も返す
    return { w48: rendered.w, h48: rendered.h };
}

// 行のテキスト色をパーツの画素から推定する
// （膨張ハローの背景色画素を混ぜると色が薄まるため、インク画素だけで平均する）
function lineColor(part, mode, line) {
    if (mode === 'inverse') return part.stats.holeColor;
    const isInk = makeInkTester(part, mode);
    let sr = 0, sg = 0, sb = 0, n = 0;
    const x0 = Math.max(0, Math.floor(line.x0)), x1 = Math.min(part.w, Math.ceil(line.x1));
    const y0 = Math.max(0, Math.floor(line.y0)), y1 = Math.min(part.h, Math.ceil(line.y1));
    for (let y = y0; y < y1; y++) {
        for (let x = x0; x < x1; x++) {
            if (!isInk(x, y)) continue;
            const o = (y * part.w + x) * 4;
            sr += part.buf[o]; sg += part.buf[o + 1]; sb += part.buf[o + 2]; n++;
        }
    }
    if (n === 0) return { r: 0, g: 0, b: 0 };
    return { r: Math.round(sr / n), g: Math.round(sg / n), b: Math.round(sb / n) };
}

// パーツのインク（文字・図の画素）判定関数を作る
//   normal  : 背景色ハローを除いた前景画素
//   inverse : 「内側の穴」と「塗りに飲み込まれた文字画素」
//             （前景がまばらな行は対象外）
function makeInkTester(part, mode) {
    const bg = part.levelBg || { r: 255, g: 255, b: 255 };
    const bar = part.stats && part.stats.top1;
    let spans = null;
    if (mode === 'inverse') {
        spans = new Array(part.h);
        for (let y = 0; y < part.h; y++) {
            const row = y * part.w * 4;
            let first = -1, last = -1, fgCount = 0;
            for (let x = 0; x < part.w; x++) {
                if (part.buf[row + x * 4 + 3] > 0) { if (first < 0) first = x; last = x; fgCount++; }
            }
            if (first >= 0 && fgCount < (last - first + 1) * 0.5) { first = -1; last = -1; }
            spans[y] = [first, last];
        }
    }
    return (x, y) => {
        if (x < 0 || y < 0 || x >= part.w || y >= part.h) return false;
        const o = (y * part.w + x) * 4;
        if (mode === 'inverse') {
            const [f, l] = spans[y];
            if (f < 0 || x <= f || x >= l) return false;
            if (part.buf[o + 3] === 0) return true;
            if (!bar) return false;
            const d = Math.abs(part.buf[o] - bar.r) + Math.abs(part.buf[o + 1] - bar.g) + Math.abs(part.buf[o + 2] - bar.b);
            return d > 150;
        }
        if (part.buf[o + 3] === 0) return false;
        const d = Math.abs(part.buf[o] - bg.r) + Math.abs(part.buf[o + 1] - bg.g) + Math.abs(part.buf[o + 2] - bg.b);
        return d > COLOR_DIST_T * 0.7;
    };
}

// 領域が「丸バッジ」（丸数字 ❶❷… など）かどうか
// バッジのインク（リング・数字・穴）は bbox の内接円の中に
// ほぼ収まる。文字の筆画は bbox の隅（円の外）にもかかる
export function isDiscLike(part, mode, box) {
    const bx0 = Math.max(0, Math.floor(box.x0)), bx1 = Math.min(part.w, Math.ceil(box.x1));
    const by0 = Math.max(0, Math.floor(box.y0)), by1 = Math.min(part.h, Math.ceil(box.y1));
    if (bx1 - bx0 < 8 || by1 - by0 < 8) return false;
    const isInk = makeInkTester(part, mode);
    // まずインクの実 bbox に切り詰める（OCR のトークン枠は行の高さ
    // いっぱいで返ることがあり、そのままでは形が測れない）
    let x0 = bx1, y0 = by1, x1 = bx0, y1 = by0;
    for (let y = by0; y < by1; y++) {
        for (let x = bx0; x < bx1; x++) {
            if (!isInk(x, y)) continue;
            if (x < x0) x0 = x; if (x > x1) x1 = x;
            if (y < y0) y0 = y; if (y > y1) y1 = y;
        }
    }
    const w = x1 - x0 + 1, h = y1 - y0 + 1;
    if (w < 8 || h < 8) return false;
    if (w / h < 0.60 || w / h > 1.7) return false; // 丸バッジはほぼ正方形
    const cx = x0 + w / 2, cy = y0 + h / 2;
    const rx = w / 2 * 1.08, ry = h / 2 * 1.08;
    let inside = 0, outside = 0;
    for (let y = y0; y <= y1; y++) {
        for (let x = x0; x <= x1; x++) {
            if (!isInk(x, y)) continue;
            const dx = (x + 0.5 - cx) / rx, dy = (y + 0.5 - cy) / ry;
            if (dx * dx + dy * dy <= 1) inside++;
            else outside++;
        }
    }
    const total = inside + outside;
    if (total < 30) return false;
    // インクがほぼ内接円に収まっていて、円の中がそれなりに濃ければバッジ
    return outside / total <= 0.10 && inside / (Math.PI * rx * ry) >= 0.30;
}

// 行の太字らしさ（文字画素の密度）を推定する
export function lineBoldness(part, mode, line) {
    let ink = 0, total = 0;
    const x0 = Math.max(0, Math.floor(line.x0)), x1 = Math.min(part.w, Math.ceil(line.x1));
    const y0 = Math.max(0, Math.floor(line.y0)), y1 = Math.min(part.h, Math.ceil(line.y1));
    const bar = part.stats && part.stats.top1;
    for (let y = y0; y < y1; y++) {
        const row = y * part.w * 4;
        let first = -1, last = -1, fgCount = 0;
        if (mode === 'inverse') {
            for (let x = 0; x < part.w; x++) {
                if (part.buf[row + x * 4 + 3] > 0) { if (first < 0) first = x; last = x; fgCount++; }
            }
            // 前景がまばらな行（塗りの上下端のアンチエイリアス）は
            // 「穴 = 文字」の判定ができないので数えない
            if (first < 0 || fgCount < (last - first + 1) * 0.5) { total += x1 - x0; continue; }
        }
        for (let x = x0; x < x1; x++) {
            const o = row + x * 4;
            let filled;
            if (mode === 'inverse') {
                filled = x > first && x < last &&
                    (part.buf[o + 3] === 0 || (bar &&
                        Math.abs(part.buf[o] - bar.r) + Math.abs(part.buf[o + 1] - bar.g) +
                        Math.abs(part.buf[o + 2] - bar.b) > 150));
            } else {
                filled = part.buf[o + 3] > 0;
            }
            if (filled) ink++;
            total++;
        }
    }
    return total > 0 ? ink / total : 0;
}

// 採用された行の範囲にあるか（少し余白を持たせて判定する）
function insideLines(lines, x, y, margin) {
    for (const l of lines) {
        if (x >= l.x0 - margin && x <= l.x1 + margin &&
            y >= l.y0 - margin && y <= l.y1 + margin) return true;
    }
    return false;
}

// 白抜き文字パーツから「採用された行の文字」だけを塗り色で消す
// （テキストボックスを上に重ねる。バッジ等の読めなかった部分は残す）
// 「穴」(透明画素) と「塗りに飲み込まれた文字画素」の両方を塗りつぶす
function fillHoles(part, lines) {
    const { r, g, b } = part.stats.top1;
    for (let y = 0; y < part.h; y++) {
        const row = y * part.w * 4;
        let first = -1, last = -1;
        for (let x = 0; x < part.w; x++) if (part.buf[row + x * 4 + 3] > 0) { if (first < 0) first = x; last = x; }
        if (first < 0) continue;
        for (let x = first + 1; x < last; x++) {
            if (!insideLines(lines, x, y, 2)) continue;
            const o = row + x * 4;
            const d = part.buf[o + 3] === 0 ? Infinity
                : Math.abs(part.buf[o] - r) + Math.abs(part.buf[o + 1] - g) + Math.abs(part.buf[o + 2] - b);
            if (d > 150) {
                part.buf[o] = r; part.buf[o + 1] = g; part.buf[o + 2] = b; part.buf[o + 3] = 255;
            }
        }
    }
}

// 通常文字パーツから「採用された行の画素」だけを透明化する
// （テキストボックスに置き換わる部分。バッジ等の読めなかった部分は
//   画像として残す）。残ったインクが無ければ画像ごと省略できるので
//   その判定結果を返す
function clearLines(part, lines) {
    const bg = part.levelBg || { r: 255, g: 255, b: 255 };
    let remain = 0;
    for (let y = 0; y < part.h; y++) {
        for (let x = 0; x < part.w; x++) {
            const o = (y * part.w + x) * 4;
            if (part.buf[o + 3] === 0) continue;
            if (insideLines(lines, x, y, 2)) {
                part.buf[o] = 0; part.buf[o + 1] = 0; part.buf[o + 2] = 0; part.buf[o + 3] = 0;
            } else {
                // 背景色のハロー以外のインクが残っているか数える
                const d = Math.abs(part.buf[o] - bg.r) + Math.abs(part.buf[o + 1] - bg.g) + Math.abs(part.buf[o + 2] - bg.b);
                if (d > COLOR_DIST_T * 0.7) remain++;
            }
        }
    }
    return remain >= 25; // true = 画像も残す必要がある
}

// 複数パーツを 1 枚の画像に縦積みして一括 OCR する
// （tesseract は起動のたびに日本語モデルを読み込むため、
//   パーツごとに実行すると非常に遅い。まとめて 1 回で認識する）
const OCR_BATCH_GAP = 64;          // パーツ間の余白 (px)。行が混ざらないよう大きめ
const OCR_BATCH_MAX_H = 12000;     // 1 バッチの最大高さ (px)

async function ocrBatch(items, tmpDir) {
    if (items.length === 0) return;
    // 縦積みの座標を決める
    let top = OCR_BATCH_GAP;
    for (const it of items) {
        it.top = top;
        top += it.h + OCR_BATCH_GAP;
    }
    const width = Math.max(...items.map(it => it.w)) + 16;
    const png = await sharp({
        create: { width, height: top, channels: 3, background: { r: 255, g: 255, b: 255 } },
    }).composite(items.map(it => ({ input: it.png, left: 8, top: it.top }))).png().toBuffer();

    const pngPath = path.join(tmpDir, `${crypto.randomUUID()}.png`);
    fs.writeFileSync(pngPath, png);
    const tsv = await runTesseract(pngPath);
    fs.unlinkSync(pngPath);
    if (!tsv) return;

    // 単語を Y 座標で各パーツに割り当て、パーツごとに行を組み立てる
    const perItem = new Map(); // item → Map(lineKey → words[])
    for (const row of tsv.split('\n')) {
        const c = row.split('\t');
        if (c.length < 12 || c[0] !== '5') continue;
        const text = c[11].trim();
        if (!text) continue;
        const bx0 = parseInt(c[6], 10), by0 = parseInt(c[7], 10);
        const bw = parseInt(c[8], 10), bh = parseInt(c[9], 10);
        const yc = by0 + bh / 2;
        const item = items.find(it => yc >= it.top && yc < it.top + it.h);
        if (!item) continue;
        if (!perItem.has(item)) perItem.set(item, new Map());
        const lineMap = perItem.get(item);
        const key = `${c[2]}_${c[3]}_${c[4]}`;
        if (!lineMap.has(key)) lineMap.set(key, []);
        // バッチ座標 → パーツローカル座標（拡大率も戻す）
        lineMap.get(key).push({
            text,
            conf: parseFloat(c[10]),
            x0: (bx0 - 8) / item.scale, y0: (by0 - item.top) / item.scale,
            x1: (bx0 + bw - 8) / item.scale, y1: (by0 + bh - item.top) / item.scale,
        });
    }

    // パーツごとに確信度ゲートを通し、通ったものに結果を付与する
    for (const it of items) {
        const lineMap = perItem.get(it);
        it.result = lineMap
            ? gateLines([...lineMap.values()], it.part.w, it.part.h,
                (wordBox) => isDiscLike(it.part, it.mode, wordBox), it.maxLineH)
            : null;
    }
}

// パーツ群に OCR をかけ、確信度の高いものにテキスト行情報を付与する
export async function ocrParts(seg) {
    if (!(await isTesseractAvailable())) return;
    const candidates = [];
    for (const part of seg.parts) {
        const mode = classifyForOcr(part, seg.workW, seg.workH);
        if (mode) candidates.push({ part, mode });
        if (candidates.length >= OCR_MAX_PARTS) break;
    }
    if (candidates.length === 0) return;

    const started = Date.now();
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'decomp_ocr_'));
    try {
        // 指定モードの候補一覧をバッチに詰めて一括 OCR する
        const runPass = async (list) => {
            const items = [];
            for (const { part, mode } of list) {
                const img = await buildOcrImage(part, mode);
                items.push({ part, mode, ...img, maxLineH: seg.workH * 0.04 });
            }
            // 高さ上限でバッチを分割する
            let batch = [], h = 0;
            const batches = [];
            for (const it of items) {
                if (h + it.h + OCR_BATCH_GAP > OCR_BATCH_MAX_H && batch.length > 0) {
                    batches.push(batch); batch = []; h = 0;
                }
                batch.push(it); h += it.h + OCR_BATCH_GAP;
            }
            if (batch.length > 0) batches.push(batch);
            for (const b of batches) {
                if (Date.now() - started > OCR_TOTAL_BUDGET_MS) break;
                await ocrBatch(b, tmpDir);
            }
            return items;
        };

        // パス1: 分類されたモードで一括認識
        const pass1 = await runPass(candidates);

        // パス2: 通常モードで読めなかった「塗りの濃いパーツ」を白抜きとして再挑戦
        // （白抜きパーツを通常モードで読むと、帯に穴が開き色も崩れるため
        //   その方向のフォールバックはしない）
        const retry = [];
        for (const it of pass1) {
            if (it.result) continue;
            if (it.mode === 'normal' && it.part.stats.holeRatio >= 0.02 &&
                it.part.stats.coreFillRatio >= 0.35 && it.part.stats.holeColor) {
                retry.push({ part: it.part, mode: 'inverse' });
            }
        }
        const pass2 = retry.length > 0 && Date.now() - started < OCR_TOTAL_BUDGET_MS
            ? await runPass(retry) : [];

        // 採用された結果を再描画照合にかけ、合格したものだけパーツに反映する
        for (const it of [...pass1, ...pass2]) {
            if (!it.result || it.part.textLines) continue;
            // 処理モードとパーツの塗り率の整合チェック
            // （白抜き帯を通常モードで処理すると表示が崩れる）
            if (it.mode === 'normal' && it.part.stats.coreFillRatio > 0.55) continue;
            if (it.mode === 'inverse' && it.part.stats.coreFillRatio < 0.50) continue;
            const lines = [];
            let allOk = true;
            for (const line of it.result) {
                const bold = lineBoldness(it.part, it.mode, line) >= 0.25;
                const v = await verifyLine(it.part, it.mode, line, bold);
                if (!v) { allOk = false; break; }
                lines.push({
                    ...line,
                    color: lineColor(it.part, it.mode, line),
                    bold,
                    hasCjk: /[⺀-鿿぀-ヿ]/.test(line.text),
                    natW48: v.w48, // フォントサイズ 48px で描画したときの自然幅
                });
            }
            // 1 行でも形が合わなければ、誤読の可能性が高いので画像のまま残す
            if (!allOk || lines.length === 0) continue;
            it.part.textLines = lines;
            it.part.ocrMode = it.mode;
            if (it.mode === 'inverse') {
                // 図形から採用行の文字だけを消し、テキストを上に重ねる
                fillHoles(it.part, lines);
                it.part.keepImage = true;
            } else {
                // 採用行の画素だけ透明化。バッジ等が残っていれば画像も出す
                it.part.keepImage = clearLines(it.part, lines);
            }
        }
    } finally {
        fs.rmSync(tmpDir, { recursive: true, force: true });
    }
}

// =========================================================
// PPTX の走査と再構築
// =========================================================

function attrInt(str, re) {
    const m = str.match(re);
    return m ? parseInt(m[1], 10) : null;
}

const escapeXml = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const hex2 = (v) => Math.max(0, Math.min(255, v)).toString(16).padStart(2, '0');
const colorHex = (c) => `${hex2(c.r)}${hex2(c.g)}${hex2(c.b)}`.toUpperCase();

// OCR された 1 行を編集可能なテキストボックス <p:sp> にする
function textboxXml(line, id, x, y, cx, cy) {
    // フォントサイズ: 行の高さから逆算（全角と欧文で文字の詰まり方が違う）
    const factor = line.hasCjk ? 0.92 : 0.74;
    let pt = (cy / 12700) * factor;
    // 長体（横に潰した文字）対策: フォントの自然幅が枠幅を超える場合は
    // 幅に収まるサイズまで縮める（照合時に測った描画幅を使う）
    if (line.natW48) {
        const ptForWidth = (cx * 48) / (12700 * line.natW48);
        pt = Math.min(pt, ptForWidth);
    }
    pt = Math.max(5, Math.min(90, pt));
    const sz = Math.round(pt * 100);
    const bold = line.bold ? ' b="1"' : '';
    return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="OCRテキスト"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>` +
        `<p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"><a:noAutofit/></a:bodyPr><a:lstStyle/>` +
        `<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="ja-JP" sz="${sz}"${bold} spc="0">` +
        `<a:solidFill><a:srgbClr val="${colorHex(line.color)}"/></a:solidFill>` +
        `<a:latin typeface="Meiryo UI"/><a:ea typeface="Meiryo UI"/></a:rPr>` +
        `<a:t>${escapeXml(line.text)}</a:t></a:r></a:p></p:txBody></p:sp>`;
}

// スライド 1 枚を処理する。分解した場合は true を返す
async function processSlide(zip, slidePath, slideW, slideH, state, options) {
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

        // 文字らしいパーツを OCR してテキスト行情報を付与する
        if (options.ocr) {
            try { await ocrParts(seg); } catch (e) { console.warn(`🔤 OCR をスキップ: ${e.message}`); }
        }

        // ピクセル座標 → EMU 座標のスケール
        const sx = extCx / seg.workW, sy = extCy / seg.workH;
        const toEmu = (p) => ({
            x: offX + Math.round(p.x0 * sx), y: offY + Math.round(p.y0 * sy),
            cx: Math.max(1, Math.round((p.x1 - p.x0) * sx)), cy: Math.max(1, Math.round((p.y1 - p.y0) * sy)),
        });

        // 既存 ID の最大値を調べて新規 ID を採番する
        let maxShapeId = 0;
        for (const m of xml.matchAll(/\bid="(\d+)"/g)) maxShapeId = Math.max(maxShapeId, parseInt(m[1], 10));
        let maxRid = 0;
        for (const m of rels.matchAll(/Id="rId(\d+)"/g)) maxRid = Math.max(maxRid, parseInt(m[1], 10));

        const newShapes = [];
        const newRels = [];
        let textboxCount = 0;
        let partNo = 0;
        for (const p of seg.parts) {
            partNo++;
            // OCR に成功したパーツは、読めた行を画像から消してテキストボックスに置き換える。
            // 読めなかった部分（バッジ・図記号など）が残る場合は画像も一緒に出す
            const emitImage = !p.textLines || p.keepImage;
            if (emitImage) {
                const png = await sharp(p.buf, { raw: { width: p.w, height: p.h, channels: 4 } })
                    .png().toBuffer();
                const mediaName = `decomp${++state.mediaSeq}.png`;
                zip.file(`ppt/media/${mediaName}`, png);
                const newRid = `rId${++maxRid}`;
                newRels.push(`<Relationship Id="${newRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${mediaName}"/>`);
                const e = toEmu({ x0: p.x, y0: p.y, x1: p.x + p.w, y1: p.y + p.h });
                const name = p.kind === 'backplate' ? '図解の背景'
                    : p.kind === 'panel' ? `図解の枠 ${partNo}` : `図解パーツ ${partNo}`;
                newShapes.push(
                    `<p:pic><p:nvPicPr><p:cNvPr id="${++maxShapeId}" name="${name}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
                    `<p:blipFill><a:blip r:embed="${newRid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
                    `<p:spPr><a:xfrm><a:off x="${e.x}" y="${e.y}"/><a:ext cx="${e.cx}" cy="${e.cy}"/></a:xfrm>` +
                    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
                );
            }
            if (p.textLines) {
                for (const line of p.textLines) {
                    const e = toEmu({
                        x0: p.x + line.x0, y0: p.y + line.y0,
                        x1: p.x + line.x1, y1: p.y + line.y1,
                    });
                    newShapes.push(textboxXml(line, ++maxShapeId, e.x, e.y, e.cx, e.cy));
                    textboxCount++;
                }
            }
        }

        // 元の 1 枚画像を、パーツ群（Z順: 背景板 → 各パーツ → テキスト）で置き換える
        xml = xml.replace(pic, newShapes.join(''));
        rels = rels.replace('</Relationships>', newRels.join('') + '</Relationships>');
        changed = true;
        console.log(`🧩 一枚絵を検出: ${slidePath} → ${seg.parts.length - 1} パーツに分解` +
            (textboxCount ? `（うち ${textboxCount} 行をテキスト化）` : ''));
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
export async function decomposeFlatImages(pptxPath, { ocr = true } = {}) {
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
            if (await processSlide(zip, slidePath, slideW, slideH, state, { ocr })) anyChanged = true;
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
