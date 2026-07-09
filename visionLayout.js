// =========================================================
// 一枚絵の Vision レイアウト解析（Claude vision）
//
// フラット化された図解（1枚のラスタ画像）を Claude の vision で
// 構造化解析し、各要素の種別・領域・テキスト・色などを持つ
// 「シーングラフ」を返す。これを decompose.js の再構築側が
// ネイティブ図形＋テキスト＋物体クロップに組み直す。
//
// 画素処理（連結成分）が「意味の分からない断片」しか作れないのに対し、
// vision LLM は「これはパネル、これは写真、これはキャプション」という
// 意味と正確な文字列を与える。両者を組み合わせるのが v3 の核心。
//
// API キー（ANTHROPIC_API_KEY）が無い／呼び出しに失敗した場合は
// null を返し、呼び出し側は従来の画素分解にフォールバックする。
// =========================================================
import { createRequire } from 'module';
const require = createRequire(import.meta.url);

const sharp = require('sharp');

// Claude vision は長辺 2576px までを高解像度で扱い、返す座標は
// 画素と 1:1 対応する。安全側で 2000px に収めてから送る
const VISION_MAX_EDGE = 2000;
const VISION_MODEL = process.env.DECOMP_VISION_MODEL || 'claude-opus-4-8';

// シーングラフの JSON スキーマ（構造化出力で形式を強制する）
// 構造化出力の制約に従い、数値レンジ制約や minLength は使わない
const SCENE_SCHEMA = {
    type: 'object',
    additionalProperties: false,
    required: ['width', 'height', 'background_color', 'elements'],
    properties: {
        width: { type: 'integer' },
        height: { type: 'integer' },
        background_color: { type: 'string' },
        elements: {
            type: 'array',
            items: {
                type: 'object',
                additionalProperties: false,
                required: ['kind', 'x', 'y', 'w', 'h', 'z'],
                properties: {
                    kind: { type: 'string', enum: ['panel', 'text', 'photo', 'icon', 'logo', 'decor'] },
                    x: { type: 'integer' },
                    y: { type: 'integer' },
                    w: { type: 'integer' },
                    h: { type: 'integer' },
                    z: { type: 'integer' },
                    text: { type: 'string' },
                    text_color: { type: 'string' },
                    bold: { type: 'boolean' },
                    align: { type: 'string', enum: ['left', 'center', 'right'] },
                    fill_color: { type: 'string' },
                    border_color: { type: 'string' },
                    rounded: { type: 'boolean' },
                },
            },
        },
    },
};

const SYSTEM_PROMPT =
    'You are a meticulous document-layout analyst. You decompose a flattened ' +
    'infographic image into a structured scene graph so it can be rebuilt as an ' +
    'editable slide (native shapes + editable text boxes + cropped images). ' +
    'You return only the structured data requested.';

function buildUserPrompt(w, h) {
    return (
        `This image is a flattened infographic, ${w}px wide and ${h}px tall. ` +
        `Decompose it into a flat list of visual elements for reconstruction as an editable slide.\n\n` +
        `Coordinates are pixels of THIS image: x,y is the top-left corner, w,h are width/height. ` +
        `Keep every box inside 0..${w} / 0..${h}.\n\n` +
        `Element kinds:\n` +
        `- "panel": a card / rounded box / colored band that acts as a background container for other elements. ` +
        `Give fill_color and border_color (hex, or "" if none) and rounded (true if the corners are rounded). ` +
        `Panels are drawn behind their contents — still list the contents as their own separate elements.\n` +
        `- "text": ONE line (or one tight run) of text. Give the exact text verbatim (preserve Japanese characters, ` +
        `punctuation, numbers, +/- and % signs), text_color (hex), bold (true/false), and align (left/center/right). ` +
        `Split multi-line captions into one element per visual line.\n` +
        `- "photo": a photograph or realistic product image.\n` +
        `- "icon": a small pictogram / line-drawing / symbol / numbered badge.\n` +
        `- "logo": a brand logo.\n` +
        `- "decor": arrows, rules, or other decoration that is not text and not a photo.\n\n` +
        `Rules:\n` +
        `- Transcribe EVERY piece of text, including titles, headings, captions, numbers, and footnotes. Accuracy of the text matters most.\n` +
        `- Make bounding boxes tight around the actual ink of each element.\n` +
        `- z is the stacking order: 0 for the backmost (panels), higher numbers for things drawn on top (text on a panel is above the panel).\n` +
        `- Colors are hex like "#1A7A3C". Sample the real color from the image.\n` +
        `- For non-text kinds, set text to "".\n`
    );
}

// Anthropic SDK クライアントを作る（キーが無ければ null）
function makeClient(apiKey) {
    const key = apiKey || process.env.ANTHROPIC_API_KEY;
    if (!key) return null;
    let AnthropicPkg;
    try {
        AnthropicPkg = require('@anthropic-ai/sdk');
    } catch (e) {
        return null; // SDK 未インストール
    }
    const Anthropic = AnthropicPkg.default || AnthropicPkg;
    return new Anthropic({ apiKey: key });
}

// vision が利用可能か（キー + SDK が揃っているか）
export function isVisionAvailable(apiKey) {
    return makeClient(apiKey) !== null;
}

// テスト用のシーン注入（API キー無しで再構築を検証するため）
// 本番では未使用。値をセットすると analyzeLayout がそれを返す
let injectedScene = null;
export function __setTestScene(scene) { injectedScene = scene; }

// 画像を Claude vision で解析し、元画像ピクセル座標のシーングラフを返す
// 失敗時は null（呼び出し側は画素分解にフォールバック）
export async function analyzeLayout(imageBuffer, { apiKey } = {}) {
    // 元画像サイズ
    const meta = await sharp(imageBuffer, { limitInputPixels: 268402689 }).metadata();
    if (!meta.width || !meta.height) return null;
    const origW = meta.width, origH = meta.height;

    // テスト注入シーンがあればそれを使う（既に元画像ピクセル座標）
    if (injectedScene) return normalizeScene(injectedScene, 1, origW, origH);

    const client = makeClient(apiKey);
    if (!client) return null;

    // 送信用に長辺 VISION_MAX_EDGE 以内へ縮小（座標系はこの送信画像基準）
    const longEdge = Math.max(origW, origH);
    const scale = longEdge > VISION_MAX_EDGE ? VISION_MAX_EDGE / longEdge : 1;
    const sendW = Math.round(origW * scale), sendH = Math.round(origH * scale);
    const jpeg = await sharp(imageBuffer, { limitInputPixels: 268402689 })
        .resize(sendW, sendH, { fit: 'fill' })
        .jpeg({ quality: 90 })
        .toBuffer();
    const b64 = jpeg.toString('base64'); // base64 は改行なし

    let raw;
    try {
        const resp = await client.messages.create({
            model: VISION_MODEL,
            max_tokens: 16000,
            thinking: { type: 'disabled' }, // 構造化出力で形式を固定するため思考は不要
            system: SYSTEM_PROMPT,
            messages: [{
                role: 'user',
                content: [
                    { type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: b64 } },
                    { type: 'text', text: buildUserPrompt(sendW, sendH) },
                ],
            }],
            output_config: { format: { type: 'json_schema', schema: SCENE_SCHEMA } },
        });
        raw = (resp.content || []).filter(b => b.type === 'text').map(b => b.text).join('');
        if (resp.stop_reason === 'refusal') return null;
    } catch (e) {
        console.warn(`🔮 vision 解析に失敗: ${e.message}`);
        return null;
    }

    let scene;
    try {
        scene = JSON.parse(raw);
    } catch (e) {
        console.warn(`🔮 vision 応答の JSON 解析に失敗: ${e.message}`);
        return null;
    }
    return normalizeScene(scene, scale, origW, origH);
}

// シーングラフを元画像ピクセル座標へスケールし、クランプ・整形する
export function normalizeScene(scene, scale, origW, origH) {
    if (!scene || !Array.isArray(scene.elements)) return null;
    const inv = scale ? 1 / scale : 1;
    const clampHex = (s) => (typeof s === 'string' && /^#?[0-9A-Fa-f]{6}$/.test(s.trim()))
        ? ('#' + s.trim().replace(/^#/, '').toUpperCase()) : '';
    const elements = [];
    for (const e of scene.elements) {
        let x = Math.round((e.x || 0) * inv);
        let y = Math.round((e.y || 0) * inv);
        let w = Math.round((e.w || 0) * inv);
        let h = Math.round((e.h || 0) * inv);
        // 画像内にクランプ
        x = Math.max(0, Math.min(origW - 1, x));
        y = Math.max(0, Math.min(origH - 1, y));
        w = Math.max(1, Math.min(origW - x, w));
        h = Math.max(1, Math.min(origH - y, h));
        const kind = ['panel', 'text', 'photo', 'icon', 'logo', 'decor'].includes(e.kind) ? e.kind : 'decor';
        elements.push({
            kind, x, y, w, h,
            z: Number.isFinite(e.z) ? e.z : 0,
            text: typeof e.text === 'string' ? e.text : '',
            textColor: clampHex(e.text_color),
            bold: !!e.bold,
            align: ['left', 'center', 'right'].includes(e.align) ? e.align : 'left',
            fill: clampHex(e.fill_color),
            border: clampHex(e.border_color),
            rounded: !!e.rounded,
        });
    }
    if (elements.length === 0) return null;
    return {
        w: origW, h: origH,
        background: clampHex(scene.background_color) || '#FFFFFF',
        elements,
    };
}
