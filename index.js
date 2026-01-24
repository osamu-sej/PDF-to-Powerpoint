
import { createRequire } from 'module';
const require = createRequire(import.meta.url);

const express = require('express');
const multer = require('multer');
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

// ★追加：パス操作のための準備
import { fileURLToPath } from 'url';
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const upload = multer({ dest: 'uploads/' });

// --- 設定などはそのまま ---
const LOG_DIR = 'secret_logs';
const ARCHIVE_DIR = path.join(LOG_DIR, 'pdf_archive');
const LOG_FILE = path.join(LOG_DIR, 'history.json');
const STANDARD_SIZES = [6, 7, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];

if (!fs.existsSync(LOG_DIR)) fs.mkdirSync(LOG_DIR);
if (!fs.existsSync(ARCHIVE_DIR)) fs.mkdirSync(ARCHIVE_DIR);

function getOneSizeSmaller(currentSizePt) {
    let closest = STANDARD_SIZES.reduce((prev, curr) => 
        (Math.abs(curr - currentSizePt) < Math.abs(prev - currentSizePt) ? curr : prev)
    );
    let index = STANDARD_SIZES.indexOf(closest);
    if (index > 0) return STANDARD_SIZES[index - 1];
    return closest;
}

function recordHistory(inputPath, originalName) {
    try {
        const now = new Date();
        const timeStr = now.toISOString().replace(/[-:T]/g, '').split('.')[0];
        const backupFilename = `${timeStr}_${originalName}`;
        const backupPath = path.join(ARCHIVE_DIR, backupFilename);
        if (fs.existsSync(inputPath)) fs.copyFileSync(inputPath, backupPath);

        const logEntry = { timestamp: now.toISOString(), original_name: originalName, archived_as: backupFilename, status: "Success" };
        let logs = [];
        if (fs.existsSync(LOG_FILE)) { try { logs = JSON.parse(fs.readFileSync(LOG_FILE, 'utf8')); } catch(e) {} }
        logs.push(logEntry);
        fs.writeFileSync(LOG_FILE, JSON.stringify(logs, null, 2));
    } catch (e) { console.error("Log Error:", e.message); }
}

// ★ここが重要！Web画面を表示する設定
// publicフォルダの中身（HTMLなど）をそのまま公開する
app.use(express.static(path.join(__dirname, 'public')));

// ★ルートURL (/) にアクセスが来たら index.html を返す
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// --- 変換API (変更なし) ---
app.post('/convert', upload.single('pdf'), async (req, res) => {
    if (!req.file) return res.status(400).send('No file uploaded.');
    
    // (中略：既存の変換コードと同じ)
    const inputPdfPath = req.file.path;
    const originalName = req.file.originalname;
    // ...以前のコードのまま...
    
    try {
        console.log(`📥 受信: ${originalName}`);
        try { execSync('fc-cache -fv', { stdio: 'ignore' }); } catch(e) {}
        const outDir = path.dirname(inputPdfPath);
        
        // 変換コマンド
        execSync(`soffice --headless --infilter="impress_pdf_import" --convert-to pptx:"Impress Office Open XML" "${inputPdfPath}" --outdir "${outDir}"`);

        // ファイル特定ロジック（簡易版：inputPdfPath + .pptx と仮定）
        // multerは拡張子なしで保存するので、LibreOfficeはそこに.pptxをつける
        const outputPptPath = inputPdfPath + '.pptx';

        if (fs.existsSync(outputPptPath)) {
            // フォント微調整
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

            // 記録
            recordHistory(inputPdfPath, originalName);

            // ダウンロード
            res.download(outputPptPath, `${originalName.replace('.pdf', '')}.pptx`, () => {
                if (fs.existsSync(inputPdfPath)) fs.unlinkSync(inputPdfPath);
                if (fs.existsSync(outputPptPath)) fs.unlinkSync(outputPptPath);
            });
        } else {
            throw new Error("Conversion failed.");
        }
    } catch (error) {
        console.error("Error:", error);
        res.status(500).send('Conversion failed. Please try again.');
        if (fs.existsSync(inputPdfPath)) fs.unlinkSync(inputPdfPath);
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 Server running on port ${PORT}`);
});