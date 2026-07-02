# PPTX Converter — PDF・HTML・Word・Excel・画像・Web ページを編集可能な PowerPoint へ

PDF をはじめ、さまざまなソースを「編集可能な」PowerPoint (PPTX) に変換する Web アプリです。
社内マニュアル・報告資料の再利用を想定しています。

## 対応している入力

| 入力 | 変換経路 |
|---|---|
| PDF | LibreOffice Impress の PDF インポート（テキスト・図形を編集可能なまま取り込み） |
| HTML / Web ページ (URL) | Chromium で 16:9 (13.333in × 7.5in) の PDF に印刷 → PDF 経路 |
| Word / Excel / テキスト (doc, docx, odt, rtf, txt, xls, xlsx, ods, csv) | LibreOffice で PDF 化 → PDF 経路 |
| PowerPoint 旧形式 (ppt, odp) | PPTX へ直接変換 |
| 画像 (png, jpg, gif, bmp, webp, svg) | LibreOffice で PDF 化 → PDF 経路 |

## 変換精度を上げるための仕組み

1. **フォント名の正規化マッピング**
   PDF に埋め込まれた PostScript 名（例: `MeiryoUI`, `UDDigiKyokashoNP-R`, `ArialMT`）を、
   PowerPoint (Windows) が認識できる正式名（`Meiryo UI`, `UD デジタル 教科書体 NP-R`, `Arial`）へ変換します。
   元のフォントの多様性はそのまま維持されます（旧来の「全フォント Meiryo UI 強制」はオプションとして残存）。
2. **フォントサイズの原寸維持**
   旧来の「一段階縮小」をやめ、既定では元のサイズを完全に維持します（縮小はオプション）。
3. **テキストの再折り返し防止**
   すべてのテキストボックスに `wrap="none"` と `noAutofit` を設定し、
   PowerPoint 側での勝手な折り返し・自動縮小によるレイアウト崩れを防ぎます。
4. **サーバー側フォントの充実 + fontconfig 代替マッピング**
   Noto CJK 全ウェイト・BIZ UD ゴシック/明朝・IPA・Liberation（Arial/Times 互換）・
   Carlito（Calibri 互換）等を導入し、`fonts-ja-substitute.conf` で Windows フォント名を
   対応付けることで、LibreOffice の文字幅計測を実フォントに近づけ、位置ズレを抑えます。
5. **ラスタ化された文字フチの除去とネイティブ輪郭への置換**
   「白文字＋黒フチ」のような装飾文字は、LibreOffice のインポートでフチ部分が
   単色ビットマップとして取り込まれ、正しいテキストに重なって表示を汚します。
   これを自動検出して除去し、代わりに PowerPoint ネイティブの文字輪郭を
   元のフチ色で付与することで、見た目を保ったまま編集可能にします。
6. **fit-to-box 補正**
   PDF 内で横方向に圧縮（長体）されたテキストは枠からはみ出しやすいため、
   枠幅と推定テキスト幅を比較し、はみ出す行にだけ文字間隔の微調整を入れて
   枠内に収めます（全角主体の行のみ・過剰圧縮はクランプ）。

## 変換オプション（画面の「詳細オプション」）

- **フォント**: 元のフォントを保持（推奨） / Meiryo UI に統一
- **文字サイズ**: 原寸を維持（推奨） / 一段階小さくする

## 動かし方

**必要なもの:** Node.js、LibreOffice (impress/draw/writer)、Chromium（HTML/URL 変換に使用）

```bash
npm install
npm run convert   # サーバー起動 (PORT 環境変数、既定 3000)
```

Docker（Render 等へのデプロイ用）:

```bash
docker build -t pptx-converter .
docker run -p 3000:3000 pptx-converter
```

## API

- `POST /convert` — multipart/form-data。フィールド: `file`（変換したいファイル）、`fontMode` (`auto`|`unify`)、`sizeMode` (`keep`|`shrink`)
- `POST /convert-url` — JSON `{ "url": "https://...", "fontMode": "auto", "sizeMode": "keep" }`
- `GET /secret-box` — 変換履歴ファイルの一覧（管理者用）
