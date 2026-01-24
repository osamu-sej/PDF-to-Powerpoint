import { createRequire } from 'module';
const require = createRequire(import.meta.url);

const express = require('express');
const multer = require('multer');
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const app = express();
// アップロードされたファイルを一時保存する設定
const upload = multer({ dest: 'uploads/' });

// 設定
const LOG_DIR = 'secret_logs';
const ARCHIVE_DIR = path.join(LOG_DIR, 'pdf_archive');
const LOG_FILE = path.join(LOG_DIR, 'history.json');
const STANDARD_SIZES = [6, 7, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];

// フォルダ初期化
if (!fs.existsSync(LOG_DIR)) fs.mkdirSync(LOG_DIR);
if (!fs.existsSync(ARCHIVE_DIR)) fs.mkdirSync(ARCHIVE_DIR);

// フォントサイズ縮小関数
function getOneSizeSmaller(currentSizePt) {
    let closest = STANDARD_SIZES.reduce((prev, curr) => 
        (Math.abs(curr - currentSizePt) < Math.abs(prev - currentSizePt) ? curr : prev)
    );
    let index = STANDARD_SIZES.indexOf(closest);
    if (index > 0) return STANDARD_SIZES[index - 1];
    return closest;
}

// 記録機能
function recordHistory(inputPath, originalName) {
    try {
        const now = new Date();
        const timeStr = now.toISOString().replace(/[-:T]/g, '').split('.')[0];
        const backupFilename = `${timeStr}_${originalName}`;
        const backupPath = path.join(ARCHIVE_DIR, backupFilename);
        
        if (fs.existsSync(inputPath)) {
            fs.copyFileSync(inputPath, backupPath);
        }

        const logEntry = {
            timestamp: now.toISOString(),
            original_name: originalName,
            archived_as: backupFilename,
            status: "Success"
        };

        let logs = [];
        if (fs.existsSync(LOG_FILE)) {
            try { logs = JSON.parse(fs.readFileSync(LOG_FILE, 'utf8')); } catch(e) {}
        }
        logs.push(logEntry);
        fs.writeFileSync(LOG_FILE, JSON.stringify(logs, null, 2));
    } catch (e) { console.error("Log Error:", e.message); }
}

// ★メインの変換API
app.post('/convert', upload.single('pdf'), async (req, res) => {
    if (!req.file) return res.status(400).send('No file uploaded.');

    const inputPdfPath = req.file.path;
    const originalName = req.file.originalname;
    const outputPptPath = inputPdfPath + '.pptx';

    try {
        console.log(`📥 受信: ${originalName}`);

        // 1. LibreOffice変換
        try { execSync('fc-cache -fv', { stdio: 'ignore' }); } catch(e) {}
        // outputPptPathと同じフォルダに出力させるための設定
        const outDir = path.dirname(inputPdfPath);
        execSync(`soffice --headless --infilter="impress_pdf_import" --convert-to pptx:"Impress Office Open XML" "${inputPdfPath}" --outdir "${outDir}"`);

        // LibreOfficeは拡張子を変えただけのファイルを作るのでパスを特定
        // 例: uploads/xxxx -> uploads/xxxx.pptx
        // ※ファイル名によっては調整が必要だが、multerのランダム名なら単純結合でOKの場合が多い
        // ここでは念のためディレクトリ内の最新PPTXを探す等の処理は省略し、標準挙動に依存

        // 2. フォント微調整
        if (fs.existsSync(outputPptPath)) {
            const data = fs.readFileSync(outputPptPath);
            const zip = await JSZip.loadAsync(data);
            const slideFiles = Object.keys(zip.files).filter(path => path.startsWith("ppt/slides/slide") && path.endsWith(".xml"));

            for (const filename of slideFiles) {
                let xmlContent = await zip.file(filename).async("string");
                xmlContent = xmlContent.replace(/sz="(\d+)"/g, (match, sizeVal) => {
                    const currentPt = parseInt(sizeVal, 10) / 100;
                    const newPt = getOneSizeSmaller(currentPt);
                    return `sz="${Math.round(newPt * 100)}"`;
                });
                zip.file(filename, xmlContent);
            }
            const content = await zip.generateAsync({ type: "nodebuffer" });
            fs.writeFileSync(outputPptPath, content);

            // 3. 記録
            recordHistory(inputPdfPath, originalName);

            // 4. ダウンロードさせる
            res.download(outputPptPath, `${originalName.replace('.pdf', '')}.pptx`, () => {
                // 送信完了後にお掃除
                if (fs.existsSync(inputPdfPath)) fs.unlinkSync(inputPdfPath);
                if (fs.existsSync(outputPptPath)) fs.unlinkSync(outputPptPath);
            });
        } else {
            throw new Error("Conversion failed, output not found.");
        }

    } catch (error) {
        console.error("Error:", error);
        res.status(500).send('Conversion failed.');
        if (fs.existsSync(inputPdfPath)) fs.unlinkSync(inputPdfPath);
    }
});

// サーバー起動（Renderなどのクラウドは PORT 環境変数を使う）
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 Server running on port ${PORT}`);
});