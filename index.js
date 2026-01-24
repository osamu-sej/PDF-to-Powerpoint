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

app.use(express.static(path.join(__dirname, 'public')));
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

app.post('/convert', upload.single('pdf'), async (req, res) => {
    if (!req.file) return res.status(400).send('No file uploaded.');
    
    const inputPdfPath = req.file.path;
    const originalName = req.file.originalname;
    const outDir = path.dirname(inputPdfPath);

    try {
        console.log(`📥 受信: ${originalName}`);
        try { execSync('fc-cache -fv', { stdio: 'ignore' }); } catch(e) {}
        
        // 1. 変換実行
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

                // (B) ★追加機能：フォントを全て "Meiryo UI" に強制置換
                // typeface="任意のフォント名" を typeface="Meiryo UI" に書き換え
                xmlContent = xmlContent.replace(/typeface="[^"]*"/g, 'typeface="Meiryo UI"');

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
        res.status(500).send('Conversion failed.');
        if (fs.existsSync(inputPdfPath)) fs.unlinkSync(inputPdfPath);
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 Server running on port ${PORT}`);
});