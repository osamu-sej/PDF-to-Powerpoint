# ベースイメージ（Node.js 20系 / Debian 12 bookworm）
# bookworm は LibreOffice が新しく PDF インポート精度が高い。
# BIZ UD ゴシック等のフォントパッケージもここから利用できる。
FROM node:20-bookworm

# 1. システムツールのインストール
#    - libreoffice        : PDF/Office → PPTX 変換エンジン
#    - chromium           : HTML/URL → PDF（高精度レンダリング）
#    - 各種フォント        : 変換時の文字幅計測の精度を上げる
#      * fonts-noto-cjk / extra : 日本語（Noto Sans/Serif CJK JP、全ウェイト）
#      * fonts-ipafont / ipaexfont : 日本語（IPA ゴシック・明朝）
#      * fonts-liberation2      : Arial / Times New Roman / Courier New 互換
#      * fonts-crosextra-carlito/caladea : Calibri / Cambria 互換
#      * fonts-noto-core        : 欧文・記号の広範なカバー
RUN apt-get update && apt-get install -y \
    libreoffice \
    chromium \
    tesseract-ocr \
    tesseract-ocr-jpn \
    fonts-noto-cjk \
    fonts-noto-cjk-extra \
    fonts-noto-core \
    fonts-ipafont \
    fonts-ipaexfont \
    fonts-morisawa-bizud-gothic \
    fonts-morisawa-bizud-mincho \
    fonts-liberation2 \
    fonts-crosextra-carlito \
    fonts-crosextra-caladea \
    fonts-dejavu \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# 2. 作業ディレクトリ作成
WORKDIR /app

# 3. 日本語フォントの代替マッピング設定
#    （Meiryo 等の Windows フォント名を Noto CJK に対応付け、
#      LibreOffice の文字幅計測を実フォントに近づける）
COPY fonts-ja-substitute.conf /etc/fonts/conf.d/65-ja-substitute.conf

# 4. フォントキャッシュ更新
RUN fc-cache -fv

# 5. パッケージ定義をコピーしてインストール
COPY package.json ./
RUN npm install

# 6. ソースコード一式をコピー
COPY . .

# 7. ログ保存用フォルダの作成
RUN mkdir -p secret_logs/pdf_archive uploads

# 8. Chromium のパス（index.js が参照）
ENV CHROMIUM_PATH=/usr/bin/chromium

# 9. コンテナ起動時のコマンド
CMD ["npm", "run", "convert"]
