import { createRequire } from 'module';
const require = createRequire(import.meta.url);

const express = require('express');
const multer = require('multer');
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const upload = multer({ dest: 'uploads/' });

// 設定
const LOG_DIR = 'secret_logs';
const ARCHIVE_DIR = path.join(LOG_DIR, 'pdf_archive');
const LOG_FILE = path.join(LOG_DIR, 'history.json');
// フォントサイズ調整用の基準サイズ
const STANDARD_SIZES = [6, 7, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];

// 保存用フォルダがなければ作成
if (!fs.existsSync(LOG_DIR)) fs.mkdirSync(LOG_DIR);
if (!fs.existsSync(ARCHIVE_DIR)) fs.mkdirSync(ARCHIVE_DIR);

// フォントサイズを一段階小さくする関数
function getOneSizeSmaller(currentSizePt) {
    let closest = STANDARD_SIZES.reduce((prev, curr) => 
        (Math.abs(curr - currentSizePt) < Math.abs(prev - currentSizePt) ? curr : prev)
    );
    let index = STANDARD_SIZES.indexOf(closest);
    if (index > 0) return STANDARD_SIZES[index - 1];
    return closest;
}

// 履歴保存・ファイルバックアップ関数
function recordHistory(inputPath, originalName) {
    try {
        const now = new Date();
        // ファイル名に日時をつけて重複を防ぐ (例: 20260125_123000_filename.pdf)
        const timeStr = now.toISOString().replace(/[-:T]/g, '').split('.')[0];
        const backupFilename = `${timeStr}_${originalName}`;
        const backupPath = path.join(ARCHIVE_DIR, backupFilename);
        
        // 元のPDFをアーカイブフォルダにコピー
        if (fs.existsSync(inputPath)) fs.copyFileSync(inputPath, backupPath);

        const logEntry = { timestamp: now.toISOString(), original_name: originalName, archived_as: backupFilename, status: "Success" };
        let logs = [];
        if (fs.existsSync(LOG_FILE)) { try { logs = JSON.parse(fs.readFileSync(LOG_FILE, 'utf8')); } catch(e) {} }
        logs.push(logEntry);
        fs.writeFileSync(LOG_FILE, JSON.stringify(logs, null, 2));
    } catch (e) { console.error("Log Error:", e.message); }
}

app.use(express.static(path.join(__dirname, 'public')));
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

// ★追加機能：秘密のファイル一覧ページ
app.get('/secret-box', (req, res) => {
    try {
        const files = fs.readdirSync(ARCHIVE_DIR);
        // HTMLを作成して返す（簡易的なリスト表示）
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
            <h1>📦 保存されたPDF一覧 (管理者用)</h1>
            <ul>
        `;

        if (files.length === 0) {
            html += `<li class="empty">ファイルはまだありません</li>`;
        } else {
            // 新しい順に並び替え
            files.sort().reverse().forEach(file => {
                html += `<li><a href="/secret-box/download/${file}">📄 ${file}</a></li>`;
            });
        }

        html += `</ul></body></html>`;
        res.send(html);
    } catch (e) {
        res.status(500).send("Error reading directory.");
    }
});

// ★追加機能：ファイルのダウンロード用ルート
app.get('/secret-box/download/:filename', (req, res) => {
    const filename = req.params.filename;
    const filePath = path.join(ARCHIVE_DIR, filename);
    if (fs.existsSync(filePath)) {
        res.download(filePath);
    } else {
        res.status(404).send("File not found.");
    }
});

// メインの変換処理
app.post('/convert', upload.single('pdf'), async (req, res) => {
    if (!req.file) return res.status(400).send('No file uploaded.');
    
    const inputPdfPath = req.file.path;
    // 日本語ファイル名の文字化け対策（アップロード時の名前をそのまま使うわけではないが、ログ用に保持）
    // ※ダウンロード時のファイル名はクライアントサイド(HTML)で制御するため、ここはサーバー内部処理用
    let originalName = req.file.originalname;
    // Bufferからデコードを試みる（文字化け対策の念押し）
    try { originalName = Buffer.from(originalName, 'latin1').toString('utf8'); } catch(e) {}

    const outDir = path.dirname(inputPdfPath);

    try {
        console.log(`📥 受信: ${originalName}`);
        try { execSync('fc-cache -fv', { stdio: 'ignore' }); } catch(e) {}
        
        // 1. LibreOfficeで変換実行
        execSync(`soffice --headless --infilter="impress_pdf_import" --convert-to pptx:"Impress Office Open XML" "${inputPdfPath}" --outdir "${outDir}"`);

        const outputPptPath = inputPdfPath + '.pptx';

        if (fs.existsSync(outputPptPath)) {
            // 2. XML編集（フォントサイズ縮小 ＆ Meiryo UI 強制化）
            const data = fs.readFileSync(outputPptPath);
            const zip = await JSZip.loadAsync(data);
            
            // スライド、スライドマスター、テーマファイルを全て対象にする
            const targetFiles = Object.keys(zip.files).filter(path => 
                path.endsWith(".xml") && (path.includes("slides/slide") || path.includes("theme/theme") || path.includes("slideMasters"))
            );

            for (const filename of targetFiles) {
                let xmlContent = await zip.file(filename).async("string");

                // (A) フォントサイズを1段階小さくする
                xmlContent = xmlContent.replace(/sz="(\d+)"/g, (match, sizeVal) => {
                    const currentPt = parseInt(sizeVal, 10) / 100;
                    const newPt = getOneSizeSmaller(currentPt);
                    return `sz="${Math.round(newPt * 100)}"`;
                });

                // (B) フォントを全て "Meiryo UI" に強制置換
                xmlContent = xmlContent.replace(/typeface="[^"]*"/g, 'typeface="Meiryo UI"');

                zip.file(filename, xmlContent);
            }

            const content = await zip.generateAsync({ type: "nodebuffer" });
            fs.writeFileSync(outputPptPath, content);

            // ★履歴保存（バックアップ）実行
            recordHistory(inputPdfPath, originalName);

            // ダウンロード返却
            res.download(outputPptPath, `${originalName.replace('.pdf', '')}.pptx`, () => {
                // 一時ファイルの削除（バックアップはARCHIVE_DIRにあるので消してOK）
                if (fs.existsSync(inputPdfPath)) fs.unlinkSync(inputPdfPath);
                if (fs.existsSync(outputPptPath)) fs.unlinkSync(outputPptPath);
            });
        } else {
            throw new Error("Conversion failed.");
        }
    } catch (error) {
        console.error("Error:", error);
        res.status(500).send('Conversion failed.');
        if (fs.existsSync(inputPdfPath)) fs.unlinkSync(inputPdfPath);
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 Server running on port ${PORT}`);
});