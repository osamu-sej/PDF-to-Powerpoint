// check.js (捜索強化版)
import fs from 'fs';
import path from 'path';

const basePath = path.join(process.cwd(), 'node_modules', 'pdf-parse');
const packageJsonPath = path.join(basePath, 'package.json');
const distPath = path.join(basePath, 'dist');

console.log("🔍 徹底調査を開始します...");

// 1. "案内図" (package.json) を見て、入口がどこか調べる
if (fs.existsSync(packageJsonPath)) {
    try {
        const pkg = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
        console.log("📄 package.json の 'main' 設定: ", pkg.main);
    } catch (e) {
        console.log("⚠️ package.json の読み取り失敗");
    }
} else {
    console.log("❌ package.json がありません");
}

// 2. "離れ" (distフォルダ) の中身を見る
if (fs.existsSync(distPath)) {
    console.log("📂 dist フォルダの中身:");
    const files = fs.readdirSync(distPath);
    files.forEach(f => console.log(" - " + f));
} else {
    console.log("❌ dist フォルダもありません... インストール自体が怪しいです");
}