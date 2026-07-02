import { createRequire } from 'module';
const require = createRequire(import.meta.url);

const express = require('express');
const multer = require('multer');
const { execFile } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const JSZip = require('jszip');
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// アプリのバージョン（package.json が唯一の定義元）
const APP_VERSION = require('./package.json').version;

const app = express();
const upload = multer({
    dest: 'uploads/',
    limits: { fileSize: 100 * 1024 * 1024 } // 100MB 上限
});
app.use(express.json());

// =========================================================
// 設定
// =========================================================
const LOG_DIR = 'secret_logs';
const ARCHIVE_DIR = path.join(LOG_DIR, 'pdf_archive');
const LOG_FILE = path.join(LOG_DIR, 'history.json');
const CONVERT_TIMEOUT_MS = 5 * 60 * 1000; // 変換タイムアウト: 5分

// フォントサイズ調整用の基準サイズ（レガシー縮小モード用）
const STANDARD_SIZES = [6, 7, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];

// 受け付ける拡張子と変換経路
//   pdf                    : そのまま Impress PDF インポート
//   html/htm               : Chromium で 16:9 PDF に印刷 → PDF 経路
//   office系 / 画像        : LibreOffice で PDF 化 → PDF 経路
const EXT_ROUTES = {
    '.pdf':  'pdf',
    '.html': 'html', '.htm': 'html',
    '.doc':  'office', '.docx': 'office', '.odt': 'office', '.rtf': 'office', '.txt': 'office',
    '.xls':  'office', '.xlsx': 'office', '.ods': 'office', '.csv': 'office',
    '.ppt':  'ppt-direct', '.odp': 'ppt-direct',
    '.png':  'office', '.jpg': 'office', '.jpeg': 'office', '.gif': 'office', '.bmp': 'office', '.webp': 'office', '.svg': 'office',
};

// 保存用フォルダがなければ作成
if (!fs.existsSync(LOG_DIR)) fs.mkdirSync(LOG_DIR);
if (!fs.existsSync(ARCHIVE_DIR)) fs.mkdirSync(ARCHIVE_DIR);

// =========================================================
// フォント名の正規化マップ
// PDF に埋め込まれた PostScript 名や別名を、PowerPoint (Windows)
// が認識できる正式なフォント名に変換する。
// キーは「小文字化 + 空白/ハイフン除去」した名前。
// =========================================================
const FONT_NAME_MAP = {
    // --- Meiryo 系 ---
    'meiryoui': 'Meiryo UI',
    'meiryouibold': 'Meiryo UI',
    'meiryouiitalic': 'Meiryo UI',
    'meiryo': 'Meiryo',
    'meiryobold': 'Meiryo',
    'meiryoitalic': 'Meiryo',
    // --- 游ゴシック / 游明朝 ---
    'yugothic': 'Yu Gothic',
    'yugothicui': 'Yu Gothic UI',
    'yugothicregular': 'Yu Gothic',
    'yugothicbold': 'Yu Gothic',
    'yugothiclight': 'Yu Gothic Light',
    'yugothicmedium': 'Yu Gothic Medium',
    'yugothib': 'Yu Gothic',
    'yugothr': 'Yu Gothic',
    'yugothl': 'Yu Gothic Light',
    'yugothm': 'Yu Gothic Medium',
    'yumincho': 'Yu Mincho',
    'yuminchodemibold': 'Yu Mincho Demibold',
    'yuminr': 'Yu Mincho',
    'yuminb': 'Yu Mincho',
    // --- MS 系 ---
    'msgothic': 'ＭＳ ゴシック',
    'mspgothic': 'ＭＳ Ｐゴシック',
    'msuigothic': 'MS UI Gothic',
    'msmincho': 'ＭＳ 明朝',
    'mspmincho': 'ＭＳ Ｐ明朝',
    // --- UD デジタル教科書体 ---
    'uddigikyokashonr': 'UD デジタル 教科書体 N-R',
    'uddigikyokashonb': 'UD デジタル 教科書体 N-B',
    'uddigikyokashonpr': 'UD デジタル 教科書体 NP-R',
    'uddigikyokashonpb': 'UD デジタル 教科書体 NP-B',
    'uddigikyokashonkr': 'UD デジタル 教科書体 NK-R',
    'uddigikyokashonkb': 'UD デジタル 教科書体 NK-B',
    // --- BIZ UD 系 ---
    'bizudgothic': 'BIZ UDゴシック',
    'bizudpgothic': 'BIZ UDPゴシック',
    'bizudmincho': 'BIZ UD明朝',
    'bizudminchomedium': 'BIZ UD明朝 Medium',
    'bizudpmincho': 'BIZ UDP明朝',
    'bizudpminchomedium': 'BIZ UDP明朝 Medium',
    // --- ヒラギノ (Mac 由来 → Windows 相当へ) ---
    'hirakakupron': 'Meiryo',
    'hirakakupro': 'Meiryo',
    'hirakakustdn': 'Meiryo',
    'hirakakustd': 'Meiryo',
    'hiraginokakugothicpron': 'Meiryo',
    'hiraginosans': 'Meiryo',
    'hiraminpron': 'Yu Mincho',
    'hiraminpro': 'Yu Mincho',
    'hiraginominchopron': 'Yu Mincho',
    'hiramarupron': 'Meiryo',
    // --- Noto / Source Han (Linux 由来 → Windows 相当へ) ---
    'notosanscjkjp': 'Meiryo UI',
    'notosansjp': 'Meiryo UI',
    'notoserifcjkjp': 'Yu Mincho',
    'notoserifjp': 'Yu Mincho',
    'sourcehansans': 'Meiryo UI',
    'sourcehansansjp': 'Meiryo UI',
    'sourcehanserif': 'Yu Mincho',
    'sourcehanserifjp': 'Yu Mincho',
    'ipagothic': 'ＭＳ ゴシック',
    'ipapgothic': 'ＭＳ Ｐゴシック',
    'ipamincho': 'ＭＳ 明朝',
    'ipapmincho': 'ＭＳ Ｐ明朝',
    'ipaexgothic': 'ＭＳ Ｐゴシック',
    'ipaexmincho': 'ＭＳ Ｐ明朝',
    'takaogothic': 'ＭＳ ゴシック',
    'takaopgothic': 'ＭＳ Ｐゴシック',
    'takaomincho': 'ＭＳ 明朝',
    // --- 欧文 PostScript 名 → 正式名 ---
    'arialmt': 'Arial',
    'arialboldmt': 'Arial',
    'arialitalicmt': 'Arial',
    'arialbolditalicmt': 'Arial',
    'arialnarrow': 'Arial Narrow',
    'helvetica': 'Arial',
    'helveticaneue': 'Arial',
    'timesnewromanpsmt': 'Times New Roman',
    'timesnewromanpsboldmt': 'Times New Roman',
    'timesnewromanpsitalicmt': 'Times New Roman',
    'timesroman': 'Times New Roman',
    'times': 'Times New Roman',
    'couriernewpsmt': 'Courier New',
    'couriernew': 'Courier New',
    'courier': 'Courier New',
    'segoeui': 'Segoe UI',
    'segoeuisemibold': 'Segoe UI Semibold',
    'segoeuilight': 'Segoe UI Light',
    'calibri': 'Calibri',
    'calibrilight': 'Calibri Light',
    'cambria': 'Cambria',
    'centurygothic': 'Century Gothic',
    'verdana': 'Verdana',
    'tahoma': 'Tahoma',
    'georgia': 'Georgia',
    'garamond': 'Garamond',
    'trebuchetms': 'Trebuchet MS',
    'candara': 'Candara',
    'consolas': 'Consolas',
    // --- Linux 系代替フォント → Windows 相当へ ---
    'liberationsans': 'Arial',
    'liberationserif': 'Times New Roman',
    'liberationmono': 'Courier New',
    'dejavusans': 'Arial',
    'dejavuserif': 'Times New Roman',
    'dejavusansmono': 'Courier New',
    'nimbussans': 'Arial',
    'nimbusroman': 'Times New Roman',
    'carlito': 'Calibri',
    'caladea': 'Cambria',
};

// フォント名から比較用キーを作る（小文字化・空白/ハイフン/カンマ除去）
function fontKey(name) {
    return name.toLowerCase().replace(/[\s\-_,]/g, '');
}

// ウェイト等のサフィックスを取り除くための正規表現
const FONT_SUFFIX_RE = /(bold|italic|oblique|regular|light|medium|semibold|demibold|black|heavy|thin|extralight|w[0-9])+$/;

// フォント名を正規化して PowerPoint が認識できる名前に変換する
function normalizeFontName(rawName) {
    if (!rawName) return rawName;
    // サブセットプレフィックス除去 (例: "ABCDEF+Meiryo" → "Meiryo")
    let name = rawName.replace(/^[A-Z]{6}\+/, '');
    const key = fontKey(name);

    // 1. 完全一致
    if (FONT_NAME_MAP[key]) return FONT_NAME_MAP[key];

    // 2. ウェイトサフィックスを除いて一致 (例: "YuGothic-Bold" → "yugothic")
    const baseKey = key.replace(FONT_SUFFIX_RE, '');
    if (baseKey && FONT_NAME_MAP[baseKey]) return FONT_NAME_MAP[baseKey];

    // 3. 不明なフォントは元の名前を維持（PC にあればそのまま使われる）
    return name;
}

// フォントサイズを一段階小さくする関数（レガシー縮小モード用）
function getOneSizeSmaller(currentSizePt) {
    let closest = STANDARD_SIZES.reduce((prev, curr) =>
        (Math.abs(curr - currentSizePt) < Math.abs(prev - currentSizePt) ? curr : prev)
    );
    let index = STANDARD_SIZES.indexOf(closest);
    if (index > 0) return STANDARD_SIZES[index - 1];
    return closest;
}

// =========================================================
// 外部コマンド実行ヘルパー（非同期・タイムアウト付き）
// execSync はサーバー全体をブロックするため使わない
// =========================================================
function run(cmd, args, options = {}) {
    return new Promise((resolve, reject) => {
        execFile(cmd, args, { timeout: CONVERT_TIMEOUT_MS, maxBuffer: 32 * 1024 * 1024, ...options },
            (error, stdout, stderr) => {
                if (error) {
                    error.stdout = stdout;
                    error.stderr = stderr;
                    reject(error);
                } else {
                    resolve({ stdout, stderr });
                }
            });
    });
}

// LibreOffice 実行（変換ごとに専用プロファイルを使い、並列実行のロック衝突を防ぐ）
async function runSoffice(extraArgs, workDir) {
    const profileDir = path.join(os.tmpdir(), `lo_profile_${crypto.randomUUID()}`);
    try {
        return await run('soffice', [
            '--headless', '--norestore', '--nolockcheck', '--nodefault',
            `-env:UserInstallation=file://${profileDir}`,
            ...extraArgs,
        ], { cwd: workDir });
    } finally {
        fs.rmSync(profileDir, { recursive: true, force: true });
    }
}

// Chromium 実行ファイルの検出（HTML → PDF 用）
function findChromium() {
    if (process.env.CHROMIUM_PATH && fs.existsSync(process.env.CHROMIUM_PATH)) {
        return process.env.CHROMIUM_PATH;
    }
    const candidates = [
        '/usr/bin/chromium', '/usr/bin/chromium-browser',
        '/usr/bin/google-chrome', '/usr/bin/google-chrome-stable',
        '/opt/pw-browsers/chromium',
    ];
    return candidates.find(p => fs.existsSync(p)) || null;
}

// =========================================================
// 入力 → PDF への変換（すべてのソースを一度 PDF に揃える）
// =========================================================

// HTML 文字列に 16:9 スライドサイズの @page 指定を注入する
function injectSlidePageStyle(html, baseUrl) {
    // 13.333in x 7.5in = PowerPoint 標準の 16:9 スライドサイズ
    const style = '<style>@page { size: 13.333in 7.5in; margin: 0; }</style>';
    const base = baseUrl ? `<base href="${baseUrl.replace(/"/g, '&quot;')}">` : '';
    if (/<head[^>]*>/i.test(html)) {
        return html.replace(/<head[^>]*>/i, m => `${m}${base}${style}`);
    }
    return `${base}${style}${html}`;
}

// HTML ファイル（またはURL取得結果）を Chromium で PDF に印刷する
async function htmlToPdf(htmlContent, outPdfPath, workDir, baseUrl) {
    const chromium = findChromium();
    if (!chromium) throw new Error('Chromium not found: HTML conversion is unavailable.');

    const tmpHtml = path.join(workDir, `page_${crypto.randomUUID()}.html`);
    fs.writeFileSync(tmpHtml, injectSlidePageStyle(htmlContent, baseUrl), 'utf8');

    try {
        await run(chromium, [
            '--headless', '--disable-gpu', '--no-sandbox', '--disable-dev-shm-usage',
            '--run-all-compositor-stages-before-draw', '--virtual-time-budget=15000',
            '--no-pdf-header-footer',
            `--print-to-pdf=${outPdfPath}`,
            `file://${tmpHtml}`,
        ]);
    } finally {
        if (fs.existsSync(tmpHtml)) fs.unlinkSync(tmpHtml);
    }
    if (!fs.existsSync(outPdfPath)) throw new Error('Chromium PDF export failed.');
}

// Office 文書・画像などを LibreOffice で PDF に変換する
async function officeToPdf(inputPath, workDir) {
    await runSoffice(['--convert-to', 'pdf', inputPath, '--outdir', workDir], workDir);
    const pdfPath = path.join(workDir, path.basename(inputPath, path.extname(inputPath)) + '.pdf');
    if (!fs.existsSync(pdfPath)) throw new Error('PDF export failed.');
    return pdfPath;
}

// PDF → 編集可能な PPTX（LibreOffice Impress PDF インポート）
async function pdfToPptx(pdfPath, workDir) {
    const { stdout, stderr } = await runSoffice([
        '--infilter=impress_pdf_import',
        '--convert-to', 'pptx:Impress Office Open XML',
        pdfPath, '--outdir', workDir,
    ], workDir);
    const pptxPath = path.join(workDir, path.basename(pdfPath, path.extname(pdfPath)) + '.pptx');
    if (!fs.existsSync(pptxPath)) {
        throw new Error(`PPTX conversion failed. soffice output: ${stdout} ${stderr}`);
    }
    return pptxPath;
}

// PPT/ODP → PPTX 直接変換
async function presentationToPptx(inputPath, workDir) {
    await runSoffice([
        '--convert-to', 'pptx:Impress Office Open XML',
        inputPath, '--outdir', workDir,
    ], workDir);
    const pptxPath = path.join(workDir, path.basename(inputPath, path.extname(inputPath)) + '.pptx');
    if (!fs.existsSync(pptxPath)) throw new Error('PPTX conversion failed.');
    return pptxPath;
}

// =========================================================
// テキスト幅の推定と fit-to-box 補正
//
// PDF 内のテキストには横方向の圧縮（長体・水平スケール）が
// かかっていることがあるが、LibreOffice の PDF インポートは
// これを再現できず、枠より広いテキストとして出力してしまう。
// そこで枠幅(cx)と推定テキスト幅を比較し、はみ出す段落にだけ
// 文字間隔(spc)のマイナス調整を入れて枠内に収める。
// =========================================================

// 1文字の幅を em 単位で概算する（フォント非依存の近似値）
function charWidthEm(ch) {
    const code = ch.codePointAt(0);
    if (code === 0x20) return 0.33;                                   // 半角スペース
    if (code < 0x0100) return 0.52;                                   // 半角英数（平均値）
    if (code >= 0xFF61 && code <= 0xFF9F) return 0.5;                 // 半角カナ
    if (code >= 0x2018 && code <= 0x201F) return 0.4;                 // 引用符
    return 1.0;                                                       // 全角（CJK・記号）
}

// XML エンティティのデコード（文字数・幅の計算用）
function decodeEntities(s) {
    return s.replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
        .replace(/&#(\d+);/g, (m, d) => String.fromCodePoint(parseInt(d, 10)))
        .replace(/&amp;/g, '&');
}

const EMU_PER_PT = 12700;
// 圧縮しすぎると読めなくなるため、1文字あたりの詰め幅に下限を設ける
const MAX_TRACKING_EM = 0.20;

// スライド XML 内の各テキストボックスに fit-to-box 補正を適用する
function applyFitToBox(xml) {
    return xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (sp) => {
        // 縦書きテキストは対象外
        if (/<a:bodyPr[^>]*\bvert="(?!horz)[^"]+"/.test(sp)) return sp;
        const extMatch = sp.match(/<a:ext cx="(\d+)" cy="\d+"\s*\/>/);
        if (!extMatch) return sp;
        const boxWidthPt = parseInt(extMatch[1], 10) / EMU_PER_PT;
        if (boxWidthPt < 10) return sp; // 幅が極端に小さい枠は対象外

        // 段落ごとに幅を推定して補正する（wrap="none" では段落 = 1行）
        return sp.replace(/<a:p>[\s\S]*?<\/a:p>/g, (para) => {
            let totalWidthPt = 0;
            let totalChars = 0;
            let cjkChars = 0;

            const runs = para.match(/<a:r>[\s\S]*?<\/a:r>/g) || [];
            for (const run of runs) {
                const szMatch = run.match(/sz="(\d+)"/);
                const sizePt = szMatch ? parseInt(szMatch[1], 10) / 100 : 18;
                const tMatch = run.match(/<a:t>([\s\S]*?)<\/a:t>/);
                if (!tMatch) continue;
                const text = decodeEntities(tMatch[1]);
                for (const ch of text) {
                    const w = charWidthEm(ch);
                    totalWidthPt += w * sizePt;
                    totalChars += 1;
                    if (w === 1.0) cjkChars += 1;
                }
            }

            // はみ出しが 1% 以内なら何もしない
            if (totalChars < 2 || totalWidthPt <= boxWidthPt * 1.01) return para;
            // 全角文字が半数未満の行は幅推定の信頼度が低いため対象外
            // （欧文の文字幅はフォントごとの差が大きい）
            if (cjkChars / totalChars < 0.5) return para;

            // 1文字あたりの詰め幅 (pt) を計算し、下限でクランプする
            const avgEmPt = totalWidthPt / totalChars;
            let perCharPt = (boxWidthPt - totalWidthPt) / totalChars;
            perCharPt = Math.max(perCharPt, -MAX_TRACKING_EM * avgEmPt);
            const spcDelta = Math.round(perCharPt * 100); // spc の単位は 1/100 pt

            // 各ランの spc に加算（無ければ挿入）する
            return para.replace(/<a:rPr\b([^>]*?)(\/?)>/g, (m, attrs, selfClose) => {
                let newAttrs;
                if (/\bspc="(-?\d+)"/.test(attrs)) {
                    newAttrs = attrs.replace(/\bspc="(-?\d+)"/, (mm, v) =>
                        `spc="${parseInt(v, 10) + spcDelta}"`);
                } else {
                    newAttrs = `${attrs} spc="${spcDelta}"`;
                }
                return `<a:rPr${newAttrs}${selfClose}>`;
            });
        });
    });
}

// =========================================================
// ラスタ化された文字装飾くずの除去
//
// PDF 内の「白フチ付き文字」などの装飾は、LibreOffice のインポートで
// フチ部分が単色のビットマップ画像として取り込まれ、正しいテキストの
// 上に重なって表示を汚してしまう。編集可能なテキスト本体は別途正しく
// 取り込まれているため、この種の画像は除去するのが最も忠実になる。
// 判定条件: 不透明部分がほぼ単色の画像 かつ テキスト枠と重なっている
// =========================================================
const sharp = require('sharp');

// 画像の不透明ピクセルがほぼ単色かどうかを判定し、その色を返す
// 戻り値: { mono: boolean, color: 'RRGGBB' | null }
async function analyzeImageColor(buffer) {
    try {
        const { data } = await sharp(buffer)
            .resize(64, 64, { fit: 'inside' })
            .ensureAlpha()
            .raw()
            .toBuffer({ resolveWithObject: true });
        const counts = new Map();
        let opaque = 0;
        for (let i = 0; i < data.length; i += 4) {
            if (data[i + 3] < 200) continue;
            opaque += 1;
            // RGB を 8 段階に量子化して色数を数える
            const key = ((data[i] >> 5) << 10) | ((data[i + 1] >> 5) << 5) | (data[i + 2] >> 5);
            counts.set(key, (counts.get(key) || 0) + 1);
        }
        if (opaque < 20) return { mono: true, color: null }; // ほぼ全部透明 → 装飾くず
        let max = 0, maxKey = 0;
        for (const [k, c] of counts) if (c > max) { max = c; maxKey = k; }
        if (max / opaque < 0.97) return { mono: false, color: null };
        // 量子化バケット中心を色として復元
        const toHex = (v) => ((v << 5) + 16).toString(16).padStart(2, '0');
        const color = `${toHex((maxKey >> 10) & 31)}${toHex((maxKey >> 5) & 31)}${toHex(maxKey & 31)}`;
        return { mono: true, color };
    } catch (e) {
        return { mono: false, color: null };
    }
}

// 図形内のテキストランに文字輪郭（アウトライン）を追加する
// （白フチ・黒フチ付き文字のフチをネイティブ表現で再現する）
function addTextOutline(sp, colorHex) {
    const lnFor = (szAttr) => {
        const szPt = szAttr ? parseInt(szAttr, 10) / 100 : 18;
        // 輪郭の太さはフォントサイズの約10%（0.75pt〜4.5pt にクランプ）
        const w = Math.max(9525, Math.min(57150, Math.round(szPt * 0.10 * 12700)));
        return `<a:ln w="${w}"><a:solidFill><a:srgbClr val="${colorHex}"/></a:solidFill></a:ln>`;
    };
    return sp.replace(/<a:rPr\b([^>]*?)(\/?)>/g, (m, attrs, selfClose, offset, full) => {
        // 既に輪郭 (a:ln) を持つランはそのまま
        if (selfClose !== '/' && full.startsWith('<a:ln', offset + m.length)) return m;
        const szMatch = attrs.match(/\bsz="(\d+)"/);
        const ln = lnFor(szMatch && szMatch[1]);
        if (selfClose === '/') return `<a:rPr${attrs}>${ln}</a:rPr>`;
        return `<a:rPr${attrs}>${ln}`;
    });
}

// スライドから文字装飾くずの画像を取り除く
async function removeRasterizedTextDecorations(zip, slidePath, xml) {
    const relsPath = slidePath.replace(/slides\/(slide\d+\.xml)$/, 'slides/_rels/$1.rels');
    const relsFile = zip.file(relsPath);
    if (!relsFile) return xml;
    const rels = await relsFile.async('string');
    const relMap = {};
    for (const m of rels.matchAll(/Id="(rId\d+)"[^>]*Target="\.\.\/media\/([^"]+)"/g)) {
        relMap[m[1]] = `ppt/media/${m[2]}`;
    }

    // テキストを持つ図形の矩形一覧
    const textShapes = [];
    for (const sp of xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || []) {
        if (!sp.includes('<a:t>')) continue;
        const off = sp.match(/<a:off x="(-?\d+)" y="(-?\d+)"/);
        const ext = sp.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
        if (off && ext) {
            textShapes.push({
                sp,
                x: parseInt(off[1], 10), y: parseInt(off[2], 10),
                w: parseInt(ext[1], 10), h: parseInt(ext[2], 10),
            });
        }
    }
    if (textShapes.length === 0) return xml;

    // 各画像を判定して置換リストを作る（media ごとに判定結果をキャッシュ）
    const mediaCache = new Map();
    const pics = xml.match(/<p:pic>[\s\S]*?<\/p:pic>/g) || [];
    const toRemove = [];
    const outlinePlan = new Map(); // sp文字列 → 輪郭色

    for (const pic of pics) {
        const off = pic.match(/<a:off x="(-?\d+)" y="(-?\d+)"/);
        const ext = pic.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
        const rid = pic.match(/r:embed="(rId\d+)"/);
        if (!off || !ext || !rid || !relMap[rid[1]]) continue;

        const px = parseInt(off[1], 10), py = parseInt(off[2], 10);
        const pw = parseInt(ext[1], 10), ph = parseInt(ext[2], 10);

        // 最も重なりの大きいテキスト図形を探す
        let best = null, bestInter = 0;
        for (const t of textShapes) {
            const ix = Math.max(0, Math.min(px + pw, t.x + t.w) - Math.max(px, t.x));
            const iy = Math.max(0, Math.min(py + ph, t.y + t.h) - Math.max(py, t.y));
            const inter = ix * iy;
            if (inter > bestInter) { bestInter = inter; best = t; }
        }
        if (!best || bestInter < 0.25 * Math.min(pw * ph, best.w * best.h)) continue;

        const mediaPath = relMap[rid[1]];
        if (!mediaCache.has(mediaPath)) {
            const mediaFile = zip.file(mediaPath);
            if (!mediaFile) { mediaCache.set(mediaPath, { mono: false, color: null }); continue; }
            const buf = await mediaFile.async('nodebuffer');
            mediaCache.set(mediaPath, await analyzeImageColor(buf));
        }
        const { mono, color } = mediaCache.get(mediaPath);
        if (!mono) continue;

        toRemove.push(pic);

        // テキストの文字色とビットマップの色が違う場合、
        // ビットマップは「文字のフチ」だったと判断し、同じ色の輪郭を文字に付ける
        if (color) {
            const fills = best.sp.match(/<a:rPr[^>]*><a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"/g) || [];
            const fillColors = fills.map(f => f.match(/val="([0-9A-Fa-f]{6})"/)[1].toLowerCase());
            const mainFill = fillColors[0];
            const similar = (a, b) => {
                const d = [0, 2, 4].reduce((s, i) =>
                    s + Math.abs(parseInt(a.slice(i, i + 2), 16) - parseInt(b.slice(i, i + 2), 16)), 0);
                return d < 96;
            };
            if (mainFill && !similar(mainFill, color.toLowerCase())) {
                outlinePlan.set(best.sp, color);
            }
        }
    }

    // 輪郭の付与 → くず画像の除去 の順に適用する
    for (const [sp, color] of outlinePlan) {
        xml = xml.replace(sp, addTextOutline(sp, color));
    }
    for (const pic of toRemove) {
        xml = xml.replace(pic, '');
    }
    return xml;
}

// =========================================================
// PPTX 後処理（品質向上の中核）
//
// fontMode:
//   'auto'  (推奨) … フォント名を正規化し PowerPoint が認識できる
//                    正式名に変換。元のフォントの多様性を維持する。
//   'unify'         … 全フォントを unifyFont に統一（旧来の動作）
// sizeMode:
//   'keep'  (推奨) … 元のフォントサイズを完全に維持
//   'shrink'        … 一段階縮小（旧来の動作）
// =========================================================
async function postProcessPptx(pptxPath, { fontMode = 'auto', unifyFont = 'Meiryo UI', sizeMode = 'keep' } = {}) {
    const data = fs.readFileSync(pptxPath);
    const zip = await JSZip.loadAsync(data);

    // スライド、レイアウト、マスター、テーマを全て対象にする
    const targetFiles = Object.keys(zip.files).filter(p =>
        p.endsWith('.xml') && (
            p.includes('slides/slide') || p.includes('slideLayouts') ||
            p.includes('slideMasters') || p.includes('theme/theme')
        )
    );

    // XML 属性用エスケープ
    const escapeAttr = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

    for (const filename of targetFiles) {
        let xml = await zip.file(filename).async('string');

        // (A) フォント名の処理
        if (fontMode === 'unify') {
            xml = xml.replace(/typeface="[^"]*"/g, `typeface="${escapeAttr(unifyFont)}"`);
        } else {
            xml = xml.replace(/typeface="([^"]*)"/g, (m, name) => {
                const decoded = name
                    .replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
                return `typeface="${escapeAttr(normalizeFontName(decoded))}"`;
            });
        }

        // (B) フォントサイズの処理
        if (sizeMode === 'shrink') {
            xml = xml.replace(/sz="(\d+)"/g, (m, sizeVal) => {
                const currentPt = parseInt(sizeVal, 10) / 100;
                const newPt = getOneSizeSmaller(currentPt);
                return `sz="${Math.round(newPt * 100)}"`;
            });
        }
        // 'keep' の場合は何もしない（原寸維持）

        // (C) スライド上のテキストボックスの再折り返しを防止する
        //     PDF インポートは行単位で正確な位置にテキストを置くため、
        //     PowerPoint 側の再折り返し・自動縮小はレイアウト崩れの元になる。
        if (filename.includes('slides/slide')) {
            xml = xml.replace(/<a:bodyPr([^>]*?)(\/?)>/g, (m, attrs, selfClose) => {
                let newAttrs = attrs;
                if (!/\bwrap=/.test(newAttrs)) newAttrs += ' wrap="none"';
                if (selfClose === '/') {
                    // 自己終了タグ → 自動調整オフを子要素として追加
                    return `<a:bodyPr${newAttrs}><a:noAutofit/></a:bodyPr>`;
                }
                return `<a:bodyPr${newAttrs}>`;
            });
            // 既存の autofit 指定を無効化
            xml = xml.replace(/<a:normAutofit[^>]*\/>/g, '<a:noAutofit/>');
            xml = xml.replace(/<a:spAutoFit\s*\/>/g, '<a:noAutofit/>');

            // (D) 枠からはみ出すテキストを文字間隔で枠内に収める
            xml = applyFitToBox(xml);

            // (E) ラスタ化された文字装飾くず（単色ビットマップ）を除去する
            xml = await removeRasterizedTextDecorations(zip, filename, xml);
        }

        zip.file(filename, xml);
    }

    // 生成した PPTX の「作成アプリケーション」情報にバージョンを記録する
    // （どのバージョンで変換されたファイルかを後から特定できるようにする）
    const appXmlFile = zip.file('docProps/app.xml');
    if (appXmlFile) {
        let appXml = await appXmlFile.async('string');
        const appName = `PPTX Converter ${APP_VERSION}`;
        if (/<Application>[^<]*<\/Application>/.test(appXml)) {
            appXml = appXml.replace(/<Application>[^<]*<\/Application>/, `<Application>${appName}</Application>`);
        } else {
            appXml = appXml.replace(/(<Properties[^>]*>)/, `$1<Application>${appName}</Application>`);
        }
        zip.file('docProps/app.xml', appXml);
    }

    const content = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    fs.writeFileSync(pptxPath, content);
}

// =========================================================
// 履歴保存・ファイルバックアップ
// =========================================================
function recordHistory(inputPath, originalName) {
    try {
        const now = new Date();
        // ファイル名に日時をつけて重複を防ぐ (例: 20260125_123000_filename.pdf)
        const timeStr = now.toISOString().replace(/[-:T]/g, '').split('.')[0];
        const safeName = path.basename(originalName).replace(/[\\/:*?"<>|]/g, '_');
        const backupFilename = `${timeStr}_${safeName}`;
        const backupPath = path.join(ARCHIVE_DIR, backupFilename);

        // 元ファイルをアーカイブフォルダにコピー
        if (fs.existsSync(inputPath)) fs.copyFileSync(inputPath, backupPath);

        const logEntry = { timestamp: now.toISOString(), original_name: originalName, archived_as: backupFilename, status: 'Success' };
        let logs = [];
        if (fs.existsSync(LOG_FILE)) { try { logs = JSON.parse(fs.readFileSync(LOG_FILE, 'utf8')); } catch (e) {} }
        logs.push(logEntry);
        fs.writeFileSync(LOG_FILE, JSON.stringify(logs, null, 2));
    } catch (e) { console.error('Log Error:', e.message); }
}

app.use(express.static(path.join(__dirname, 'public')));
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

// バージョン情報 API（画面のバッジ表示と、デプロイ確認に使う）
app.get('/version', (req, res) => {
    res.json({ version: APP_VERSION });
});

// ★秘密のファイル一覧ページ
app.get('/secret-box', (req, res) => {
    try {
        const files = fs.readdirSync(ARCHIVE_DIR);
        let html = `
        <!DOCTYPE html>
        <html lang="ja">
        <head>
            <meta charset="UTF-8">
            <title>Secret Box</title>
            <style>
                body { font-family: sans-serif; padding: 20px; max-width: 800px; margin: 0 auto; background: #f0f0f0; }
                h1 { color: #333; }
                ul { list-style: none; padding: 0; }
                li { background: white; margin: 10px 0; padding: 15px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
                a { text-decoration: none; color: #2563eb; font-weight: bold; }
                a:hover { text-decoration: underline; }
                .empty { color: #888; }
            </style>
        </head>
        <body>
            <h1>📦 保存されたファイル一覧 (管理者用)</h1>
            <ul>
        `;

        if (files.length === 0) {
            html += `<li class="empty">ファイルはまだありません</li>`;
        } else {
            files.sort().reverse().forEach(file => {
                html += `<li><a href="/secret-box/download/${encodeURIComponent(file)}">📄 ${file}</a></li>`;
            });
        }

        html += `</ul></body></html>`;
        res.send(html);
    } catch (e) {
        res.status(500).send('Error reading directory.');
    }
});

// ★ファイルのダウンロード用ルート（パストラバーサル対策済み）
app.get('/secret-box/download/:filename', (req, res) => {
    const filename = path.basename(req.params.filename);
    const filePath = path.join(ARCHIVE_DIR, filename);
    if (fs.existsSync(filePath)) {
        res.download(filePath);
    } else {
        res.status(404).send('File not found.');
    }
});

// =========================================================
// 変換パイプライン本体：任意ソース → PDF → 編集可能 PPTX
// =========================================================
async function convertToPptx({ route, inputPath, workDir, options }) {
    let pptxPath;

    if (route === 'pdf') {
        pptxPath = await pdfToPptx(inputPath, workDir);
    } else if (route === 'html') {
        const html = fs.readFileSync(inputPath, 'utf8');
        const pdfPath = path.join(workDir, `html_${crypto.randomUUID()}.pdf`);
        await htmlToPdf(html, pdfPath, workDir);
        pptxPath = await pdfToPptx(pdfPath, workDir);
        fs.unlinkSync(pdfPath);
    } else if (route === 'ppt-direct') {
        pptxPath = await presentationToPptx(inputPath, workDir);
    } else { // office (Word / Excel / 画像 など)
        const pdfPath = await officeToPdf(inputPath, workDir);
        pptxPath = await pdfToPptx(pdfPath, workDir);
        fs.unlinkSync(pdfPath);
    }

    await postProcessPptx(pptxPath, options);
    return pptxPath;
}

// リクエストから変換オプションを取り出す
function parseOptions(body = {}) {
    return {
        fontMode: body.fontMode === 'unify' ? 'unify' : 'auto',
        unifyFont: typeof body.unifyFont === 'string' && body.unifyFont.trim() ? body.unifyFont.trim() : 'Meiryo UI',
        sizeMode: body.sizeMode === 'shrink' ? 'shrink' : 'keep',
    };
}

// メインの変換処理（ファイルアップロード）
// フィールド名は 'file'（旧クライアント互換のため 'pdf' も受け付ける）
app.post('/convert', upload.any(), async (req, res) => {
    const file = (req.files || []).find(f => f.fieldname === 'file' || f.fieldname === 'pdf');
    if (!file) return res.status(400).send('No file uploaded.');

    const inputPath = path.resolve(file.path);
    // 日本語ファイル名の文字化け対策
    let originalName = file.originalname;
    try { originalName = Buffer.from(originalName, 'latin1').toString('utf8'); } catch (e) {}

    const ext = path.extname(originalName).toLowerCase();
    const route = EXT_ROUTES[ext];
    if (!route) {
        fs.unlinkSync(inputPath);
        return res.status(400).send(`Unsupported file type: ${ext}`);
    }

    // LibreOffice が拡張子で形式判別できるようにリネーム
    const workDir = path.dirname(inputPath);
    const renamedInput = `${inputPath}${ext}`;
    fs.renameSync(inputPath, renamedInput);

    const cleanup = (pptxPath) => {
        for (const p of [renamedInput, pptxPath]) {
            try { if (p && fs.existsSync(p)) fs.unlinkSync(p); } catch (e) {}
        }
    };

    try {
        console.log(`📥 受信: ${originalName} (${route})`);
        const options = parseOptions(req.body);
        const pptxPath = await convertToPptx({ route, inputPath: renamedInput, workDir, options });

        // ★履歴保存（バックアップ）実行
        recordHistory(renamedInput, originalName);

        const downloadName = `${originalName.replace(/\.[^.]+$/, '')}.pptx`;
        res.download(pptxPath, downloadName, () => cleanup(pptxPath));
    } catch (error) {
        console.error('Error:', error.message, error.stderr || '');
        res.status(500).send('Conversion failed.');
        cleanup();
    }
});

// URL → PPTX 変換（Web ページをスライド化）
app.post('/convert-url', async (req, res) => {
    const { url } = req.body || {};
    if (!url || !/^https?:\/\//i.test(url)) {
        return res.status(400).send('Valid http(s) URL is required.');
    }

    const workDir = path.resolve('uploads');
    if (!fs.existsSync(workDir)) fs.mkdirSync(workDir);
    const pdfPath = path.join(workDir, `url_${crypto.randomUUID()}.pdf`);
    let pptxPath;

    try {
        console.log(`📥 URL受信: ${url}`);
        // ページ本体を取得し、16:9 の @page スタイルと <base> を注入して印刷する
        const response = await fetch(url, { redirect: 'follow', signal: AbortSignal.timeout(30000) });
        if (!response.ok) throw new Error(`Fetch failed: ${response.status}`);
        const html = await response.text();
        await htmlToPdf(html, pdfPath, workDir, response.url || url);

        pptxPath = await pdfToPptx(pdfPath, workDir);
        await postProcessPptx(pptxPath, parseOptions(req.body));

        recordHistory(pdfPath, `${url}.pdf`);

        // URL からダウンロードファイル名を生成
        let host = 'webpage';
        try { host = new URL(url).hostname.replace(/[^\w.-]/g, '_'); } catch (e) {}
        res.download(pptxPath, `${host}.pptx`, () => {
            for (const p of [pdfPath, pptxPath]) {
                try { if (fs.existsSync(p)) fs.unlinkSync(p); } catch (e) {}
            }
        });
    } catch (error) {
        console.error('Error:', error.message, error.stderr || '');
        res.status(500).send('Conversion failed.');
        for (const p of [pdfPath, pptxPath]) {
            try { if (p && fs.existsSync(p)) fs.unlinkSync(p); } catch (e) {}
        }
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 PPTX Converter v${APP_VERSION} running on port ${PORT}`);
});
