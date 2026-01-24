# ベースイメージ（Node.js 20系）
FROM node:20-bullseye

# 1. システムツール（LibreOfficeと日本語フォント）のインストール
# ※Reactなどは関係なく、このOSレベルの設定が変換には必須です
RUN apt-get update && apt-get install -y \
    libreoffice \
    fonts-noto-cjk \
    fonts-ipafont \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# 2. 作業ディレクトリ作成
WORKDIR /app

# 3. フォントキャッシュ更新
RUN fc-cache -fv

# 4. パッケージ定義をコピーしてインストール
# ※ここでReactなどの依存関係も全部インストールされますが、問題ありません
COPY package.json ./
RUN npm install

# 5. ソースコード一式をコピー
# （index.js だけでなく、プロジェクト全体をコピーします）
COPY . .

# 6. ログ保存用フォルダの作成
RUN mkdir -p secret_logs/pdf_archive

# 7. コンテナ起動時のコマンド
# package.jsonに追加した "convert" コマンドを実行します
CMD ["npm", "run", "convert"]